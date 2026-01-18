using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO.Pipes;
using System.Text;
using ExchangeAdmin.Contracts;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Worker.Operations;
using ExchangeAdmin.Worker.PowerShell;

namespace ExchangeAdmin.Worker.Ipc;

/// <summary>
/// Server IPC Named Pipes per il worker.
/// Implementa framing robusto, limiti difensivi, e gestione cancel idempotente.
/// </summary>
public sealed class IpcServer : IDisposable
{
    private readonly PowerShellEngine _psEngine;
    private readonly OperationDispatcher _dispatcher;
    private readonly DateTime _startTime = DateTime.UtcNow;

    private NamedPipeServerStream? _requestPipe;
    private NamedPipeServerStream? _eventPipe;
    private StreamReader? _requestReader;
    private StreamWriter? _requestWriter;
    private StreamWriter? _eventWriter;

    private CancellationTokenSource? _serverCts;
    private Task? _requestLoopTask;

    private readonly ConcurrentDictionary<string, CancellationTokenSource> _activeOperations = new();
    private readonly ConcurrentDictionary<string, int> _eventCounts = new();
    private readonly SemaphoreSlim _eventWriteLock = new(1, 1);

    private volatile bool _isRunning;
    private volatile bool _isDisposing;
    private string? _clientId;

    /// <summary>
    /// Crea un nuovo server IPC.
    /// </summary>
    /// <param name="psEngine">Engine PowerShell per esecuzione comandi.</param>
    public IpcServer(PowerShellEngine psEngine)
    {
        _psEngine = psEngine;
        _dispatcher = new OperationDispatcher(psEngine, SendEventAsync);
    }

    /// <summary>
    /// Avvia il server IPC e attende connessione client.
    /// </summary>
    public async Task StartAsync()
    {
        _serverCts = new CancellationTokenSource();

        // Crea pipe principale per request/response
        _requestPipe = new NamedPipeServerStream(
            IpcConstants.PipeName,
            PipeDirection.InOut,
            1,
            PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous,
            IpcConstants.PipeBufferSize,
            IpcConstants.PipeBufferSize);

        // Crea pipe eventi
        _eventPipe = new NamedPipeServerStream(
            IpcConstants.EventPipeName,
            PipeDirection.Out,
            1,
            PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous,
            IpcConstants.PipeBufferSize,
            0);

        Console.WriteLine($"[IPC] Waiting for client connection on {IpcConstants.PipeName}...");

        // Attendi connessione su entrambe le pipe
        await Task.WhenAll(
            _requestPipe.WaitForConnectionAsync(_serverCts.Token),
            _eventPipe.WaitForConnectionAsync(_serverCts.Token)).ConfigureAwait(false);

        Console.WriteLine("[IPC] Client connected");

        // Crea readers/writers SENZA AutoFlush per evitare flush prematuro
        _requestReader = new StreamReader(_requestPipe, Encoding.UTF8, leaveOpen: true);
        _requestWriter = new StreamWriter(_requestPipe, Encoding.UTF8, leaveOpen: true);
        _eventWriter = new StreamWriter(_eventPipe, Encoding.UTF8, leaveOpen: true);

        Console.WriteLine("[IPC] Readers/writers created successfully");

        _isRunning = true;

        // Avvia loop richieste
        _requestLoopTask = RequestLoopAsync(_serverCts.Token);
    }

    /// <summary>
    /// Ferma il server IPC in modo pulito.
    /// </summary>
    public async Task StopAsync()
    {
        _isRunning = false;
        _serverCts?.Cancel();

        // Cancella tutte le operazioni attive
        foreach (var kvp in _activeOperations)
        {
            try
            {
                kvp.Value.Cancel();
            }
            catch (ObjectDisposedException)
            {
                // Già disposed
            }
        }
        _activeOperations.Clear();
        _eventCounts.Clear();

        if (_requestLoopTask != null)
        {
            try
            {
                await Task.WhenAny(_requestLoopTask, Task.Delay(5000)).ConfigureAwait(false);
            }
            catch
            {
                // Ignora errori durante shutdown
            }
        }

        Cleanup();
    }

    private async Task RequestLoopAsync(CancellationToken cancellationToken)
    {
        try
        {
            while (!cancellationToken.IsCancellationRequested && _isRunning && _requestReader != null)
            {
                string? line;
                try
                {
                    line = await _requestReader.ReadLineAsync(cancellationToken).ConfigureAwait(false);
                }
                catch (IOException ex)
                {
                    Console.WriteLine($"[IPC] RequestLoop IOException: {ex.Message}");
                    break;
                }

                if (line == null)
                {
                    Console.WriteLine("[IPC] Client disconnected (null read)");
                    break;
                }

                // Validazione dimensione messaggio
                if (!IpcConstants.IsValidMessageSize(line.Length))
                {
                    Console.WriteLine($"[IPC] Message too large ({line.Length} bytes), rejecting");
                    continue;
                }

                IpcMessage? message;
                try
                {
                    message = JsonMessageSerializer.DeserializeMessage(line);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[IPC] JSON parse error: {ex.Message}");
                    continue;
                }

                if (message == null)
                {
                    Console.WriteLine($"[IPC] Invalid message received (unknown type)");
                    continue;
                }

                // Gestisci messaggio in background (non bloccare il loop)
                _ = HandleMessageAsync(message, cancellationToken);
            }
        }
        catch (OperationCanceledException)
        {
            // Shutdown normale
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[IPC] Request loop fatal error: {ex.Message}");
        }
        finally
        {
            Console.WriteLine("[IPC] Request loop terminated");
        }
    }

    private async Task HandleMessageAsync(IpcMessage message, CancellationToken cancellationToken)
    {
        try
        {
            switch (message)
            {
                case HandshakeRequest handshake:
                    await HandleHandshakeAsync(handshake).ConfigureAwait(false);
                    break;

                case RequestEnvelope request:
                    await HandleRequestAsync(request, cancellationToken).ConfigureAwait(false);
                    break;

                case CancelRequest cancel:
                    HandleCancel(cancel);
                    break;

                case HeartbeatPing ping:
                    await HandleHeartbeatAsync(ping).ConfigureAwait(false);
                    break;

                default:
                    Console.WriteLine($"[IPC] Unknown message type: {message.Type}");
                    break;
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[IPC] Error handling message {message.Type}: {ex.Message}");
        }
    }

    private async Task HandleHandshakeAsync(HandshakeRequest request)
    {
        Console.WriteLine($"[IPC] Handshake request from client {request.ClientId}");
        Console.WriteLine($"[IPC] Client contracts version: {request.ContractsVersion}");

        _clientId = request.ClientId;

        var isCompatible = ContractVersion.IsCompatible(request.ContractsVersion);

        var response = new HandshakeResponse
        {
            Success = isCompatible,
            ContractsVersion = ContractVersion.Version,
            WorkerVersion = "1.0.0",
            IsModuleAvailable = _psEngine.IsModuleAvailable,
            PowerShellVersion = _psEngine.PowerShellVersion,
            ErrorMessage = isCompatible ? null : $"Incompatible contracts version. Worker: {ContractVersion.Version}, Client: {request.ContractsVersion}"
        };

        Console.WriteLine($"[IPC] Sending handshake response...");
        await SendResponseRawAsync(response).ConfigureAwait(false);
        Console.WriteLine($"[IPC] Handshake response sent successfully");

        Console.WriteLine($"[IPC] Handshake completed. Compatible: {isCompatible}");
    }

    private async Task HandleRequestAsync(RequestEnvelope request, CancellationToken serverCancellation)
    {
        Console.WriteLine($"[IPC] Request: {request.Operation} (correlation: {request.CorrelationId})");

        // Inizializza contatore eventi per questa request
        _eventCounts[request.CorrelationId] = 0;

        // Crea CancellationToken per questa operazione
        CancellationTokenSource? operationCts = null;
        try
        {
            operationCts = CancellationTokenSource.CreateLinkedTokenSource(serverCancellation);
            _activeOperations[request.CorrelationId] = operationCts;

            var response = await _dispatcher.DispatchAsync(request, operationCts.Token).ConfigureAwait(false);
            await SendResponseRawAsync(response).ConfigureAwait(false);
        }
        catch (ObjectDisposedException)
        {
            // Server in shutdown
            return;
        }
        catch (OperationCanceledException)
        {
            var response = new ResponseEnvelope
            {
                CorrelationId = request.CorrelationId,
                Success = false,
                WasCancelled = true
            };
            await SendResponseRawAsync(response).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[IPC] Operation error: {ex.Message}");

            var response = new ResponseEnvelope
            {
                CorrelationId = request.CorrelationId,
                Success = false,
                Error = new NormalizedErrorDto
                {
                    Code = ErrorCode.Unknown,
                    Message = ex.Message,
                    Details = ex.StackTrace,
                    IsTransient = false
                }
            };
            await SendResponseRawAsync(response).ConfigureAwait(false);
        }
        finally
        {
            _activeOperations.TryRemove(request.CorrelationId, out _);
            _eventCounts.TryRemove(request.CorrelationId, out _);

            if (operationCts != null)
            {
                try
                {
                    operationCts.Dispose();
                }
                catch
                {
                    // Ignora errori dispose
                }

                operationCts = null;
            }
        }
    }

    /// <summary>
    /// Gestisce richiesta di cancellazione. Idempotente - non fallisce se operazione già completata.
    /// </summary>
    private void HandleCancel(CancelRequest cancel)
    {
        Console.WriteLine($"[IPC] Cancel request for correlation: {cancel.CorrelationId}");

        if (_activeOperations.TryGetValue(cancel.CorrelationId, out var cts))
        {
            try
            {
                if (!cts.IsCancellationRequested)
                {
                    cts.Cancel();
                    Console.WriteLine($"[IPC] Cancellation signaled for: {cancel.CorrelationId}");
                }
                else
                {
                    Console.WriteLine($"[IPC] Already cancelled: {cancel.CorrelationId}");
                }
            }
            catch (ObjectDisposedException)
            {
                // Operazione già completata e CTS disposed
                Console.WriteLine($"[IPC] Operation already completed: {cancel.CorrelationId}");
            }
        }
        else
        {
            // Operazione non trovata - già completata o mai esistita
            Console.WriteLine($"[IPC] Cancel ignored - operation not found: {cancel.CorrelationId}");
        }
    }

    private async Task HandleHeartbeatAsync(HeartbeatPing ping)
    {
        var pong = new HeartbeatPong
        {
            Sequence = ping.Sequence,
            WorkerUptime = DateTime.UtcNow - _startTime,
            ActiveOperations = _activeOperations.Count
        };

        await SendResponseRawAsync(pong).ConfigureAwait(false);
    }

    private async Task SendResponseRawAsync<T>(T message) where T : IpcMessage
    {
        if (_requestWriter == null || _isDisposing)
        {
            Console.WriteLine("[IPC] Cannot send response - writer is null or disposing");
            return;
        }

        try
        {
            var json = JsonMessageSerializer.Serialize(message);

            // Verifica dimensione prima di inviare
            if (!IpcConstants.IsValidMessageSize(json.Length))
            {
                Console.Error.WriteLine($"[IPC] Response too large ({json.Length} bytes), dropping");
                return;
            }

            Console.WriteLine($"[IPC] Sending response (length: {json.Length} bytes, type: {typeof(T).Name})");
            await _requestWriter.WriteLineAsync(json).ConfigureAwait(false);
            await _requestWriter.FlushAsync().ConfigureAwait(false);
            Console.WriteLine($"[IPC] Response sent successfully");
        }
        catch (IOException ex)
        {
            Console.Error.WriteLine($"[IPC] Failed to send response: {ex.Message}");
        }
        catch (ObjectDisposedException)
        {
            // Pipe già chiusa
        }
    }

    private async Task SendEventAsync(EventEnvelope evt)
    {
        if (_eventWriter == null || _isDisposing)
        {
            return;
        }

        // Verifica limite eventi per request
        if (_eventCounts.TryGetValue(evt.CorrelationId, out var count))
        {
            if (!IpcConstants.IsEventCountWithinLimit(count))
            {
                Debug.WriteLine($"[IPC] Max events reached for {evt.CorrelationId}, dropping event");
                return;
            }
            _eventCounts[evt.CorrelationId] = count + 1;
        }

        await _eventWriteLock.WaitAsync().ConfigureAwait(false);
        try
        {
            if (_eventWriter == null || _isDisposing)
            {
                return;
            }

            var json = JsonMessageSerializer.Serialize(evt);

            // Verifica dimensione prima di inviare
            if (!IpcConstants.IsValidMessageSize(json.Length))
            {
                Console.Error.WriteLine($"[IPC] Event too large ({json.Length} bytes), dropping");
                return;
            }

            await _eventWriter.WriteLineAsync(json).ConfigureAwait(false);
            await _eventWriter.FlushAsync().ConfigureAwait(false);
        }
        catch (IOException ex)
        {
            Debug.WriteLine($"[IPC] Failed to send event: {ex.Message}");
        }
        catch (ObjectDisposedException)
        {
            // Pipe già chiusa
        }
        finally
        {
            _eventWriteLock.Release();
        }
    }

    private void Cleanup()
    {
        try { _requestReader?.Dispose(); } catch { }
        try { _requestWriter?.Dispose(); } catch { }
        try { _eventWriter?.Dispose(); } catch { }
        try { _requestPipe?.Dispose(); } catch { }
        try { _eventPipe?.Dispose(); } catch { }

        _requestReader = null;
        _requestWriter = null;
        _eventWriter = null;
        _requestPipe = null;
        _eventPipe = null;
    }

    /// <summary>
    /// Rilascia le risorse.
    /// </summary>
    public void Dispose()
    {
        if (_isDisposing)
        {
            return;
        }

        _isDisposing = true;

        _serverCts?.Cancel();

        // Cancella tutte le operazioni attive
        foreach (var kvp in _activeOperations)
        {
            try
            {
                kvp.Value.Cancel();
                kvp.Value.Dispose();
            }
            catch
            {
                // Ignora
            }
        }
        _activeOperations.Clear();

        _serverCts?.Dispose();
        _eventWriteLock.Dispose();
        Cleanup();
    }
}
