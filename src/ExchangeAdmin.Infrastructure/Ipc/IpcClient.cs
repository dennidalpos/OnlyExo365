using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO.Pipes;
using System.Text;
using ExchangeAdmin.Contracts;
using ExchangeAdmin.Contracts.Messages;

namespace ExchangeAdmin.Infrastructure.Ipc;

/// <summary>
/// Client IPC per comunicazione con il worker via Named Pipes.
/// Implementa framing robusto, protezione da messaggi parziali, e gestione disconnessione.
/// </summary>
public class IpcClient : IAsyncDisposable
{
    private readonly string _pipeName;
    private readonly string _eventPipeName;
    private NamedPipeClientStream? _requestPipe;
    private NamedPipeClientStream? _eventPipe;
    private StreamReader? _requestReader;
    private StreamWriter? _requestWriter;
    private StreamReader? _eventReader;

    private readonly ConcurrentDictionary<string, TaskCompletionSource<ResponseEnvelope>> _pendingRequests = new();
    private readonly ConcurrentDictionary<string, Action<EventEnvelope>> _eventHandlers = new();
    private readonly ConcurrentDictionary<string, int> _eventCounts = new();

    private CancellationTokenSource? _eventLoopCts;
    private Task? _eventLoopTask;
    private Task? _responseLoopTask;

    private volatile bool _isConnected;
    private volatile bool _isDisposing;
    private readonly object _stateLock = new();
    private readonly SemaphoreSlim _sendLock = new(1, 1);
    private readonly StringBuilder _partialMessageBuffer = new();

    /// <summary>
    /// Evento di cambio stato connessione.
    /// </summary>
    public event EventHandler<WorkerConnectionState>? ConnectionStateChanged;

    /// <summary>
    /// Evento ricevuto dal worker (globale).
    /// </summary>
    public event EventHandler<EventEnvelope>? EventReceived;

    /// <summary>
    /// Indica se il client è connesso al worker.
    /// </summary>
    public bool IsConnected => _isConnected;

    /// <summary>
    /// Crea un nuovo client IPC.
    /// </summary>
    /// <param name="pipeName">Nome pipe principale (null = default).</param>
    /// <param name="eventPipeName">Nome pipe eventi (null = default).</param>
    public IpcClient(string? pipeName = null, string? eventPipeName = null)
    {
        _pipeName = pipeName ?? IpcConstants.PipeName;
        _eventPipeName = eventPipeName ?? IpcConstants.EventPipeName;
    }

    /// <summary>
    /// Connette al worker e esegue handshake.
    /// </summary>
    /// <exception cref="InvalidOperationException">Se handshake fallisce o versioni incompatibili.</exception>
    /// <exception cref="TimeoutException">Se connessione o handshake vanno in timeout.</exception>
    public async Task<HandshakeResponse> ConnectAsync(CancellationToken cancellationToken = default)
    {
        ThrowIfDisposing();

        try
        {
            // Crea entrambe le pipe
            _requestPipe = new NamedPipeClientStream(".", _pipeName, PipeDirection.InOut, PipeOptions.Asynchronous);
            _eventPipe = new NamedPipeClientStream(".", _eventPipeName, PipeDirection.In, PipeOptions.Asynchronous);

            // Connetti entrambe le pipe PRIMA di creare readers/writers
            // Questo evita race condition con il server che aspetta entrambe le connessioni
            await Task.WhenAll(
                _requestPipe.ConnectAsync(IpcConstants.ConnectionTimeoutMs, cancellationToken),
                _eventPipe.ConnectAsync(IpcConstants.ConnectionTimeoutMs, cancellationToken)).ConfigureAwait(false);

            // Ora crea readers/writers DOPO che entrambe le pipe sono connesse
            // NON usare AutoFlush per evitare flush prematuro durante la creazione
            _requestReader = new StreamReader(_requestPipe, Encoding.UTF8, leaveOpen: true);
            _requestWriter = new StreamWriter(_requestPipe, Encoding.UTF8, leaveOpen: true);
            _eventReader = new StreamReader(_eventPipe, Encoding.UTF8, leaveOpen: true);

            // Handshake con timeout dedicato
            var handshakeRequest = new HandshakeRequest();
            await SendRawAsync(handshakeRequest, cancellationToken).ConfigureAwait(false);

            using var handshakeCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            handshakeCts.CancelAfter(IpcConstants.HandshakeTimeoutMs);

            var responseJson = await ReadLineWithTimeoutAsync(_requestReader, handshakeCts.Token).ConfigureAwait(false);

            if (string.IsNullOrEmpty(responseJson))
            {
                throw new InvalidOperationException("Empty handshake response from worker");
            }

            // Valida dimensione messaggio
            if (!IpcConstants.IsValidMessageSize(responseJson.Length))
            {
                throw new InvalidOperationException($"Handshake response exceeds maximum size ({responseJson.Length} bytes)");
            }

            var response = JsonMessageSerializer.Deserialize<HandshakeResponse>(responseJson);
            if (response == null)
            {
                throw new InvalidOperationException("Invalid handshake response from worker (deserialization failed)");
            }

            if (!response.Success)
            {
                throw new InvalidOperationException($"Handshake failed: {response.ErrorMessage}");
            }

            // Verifica compatibilità versione
            if (!ContractVersion.IsCompatible(response.ContractsVersion))
            {
                throw new InvalidOperationException(
                    $"Contract version mismatch. Client: {ContractVersion.Version}, Worker: {response.ContractsVersion}");
            }

            lock (_stateLock)
            {
                _isConnected = true;
            }

            // Avvia loop eventi (usano il proprio CTS, indipendente dal token di connessione)
            _eventLoopCts = new CancellationTokenSource();
            _eventLoopTask = Task.Run(() => EventLoopAsync(_eventLoopCts.Token), CancellationToken.None);
            _responseLoopTask = Task.Run(() => ResponseLoopAsync(_eventLoopCts.Token), CancellationToken.None);

            ConnectionStateChanged?.Invoke(this, WorkerConnectionState.Connected);

            return response;
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            // Cleanup in caso di errore durante connessione
            await DisposeInternalAsync().ConfigureAwait(false);
            throw;
        }
    }

    /// <summary>
    /// Invia una richiesta e attende la risposta.
    /// </summary>
    /// <param name="request">Request envelope.</param>
    /// <param name="eventHandler">Handler opzionale per eventi streaming.</param>
    /// <param name="cancellationToken">Token di cancellazione.</param>
    /// <returns>Response envelope dal worker.</returns>
    /// <exception cref="InvalidOperationException">Se non connesso.</exception>
    /// <exception cref="OperationCanceledException">Se cancellato o timeout.</exception>
    public async Task<ResponseEnvelope> SendRequestAsync(
        RequestEnvelope request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
    {
        ThrowIfDisposing();

        if (!_isConnected)
        {
            throw new InvalidOperationException("Not connected to worker");
        }

        var tcs = new TaskCompletionSource<ResponseEnvelope>(TaskCreationOptions.RunContinuationsAsynchronously);

        // Registra handler eventi per questa correlazione
        if (eventHandler != null)
        {
            _eventHandlers[request.CorrelationId] = eventHandler;
            _eventCounts[request.CorrelationId] = 0;
        }

        // Registra pending request
        _pendingRequests[request.CorrelationId] = tcs;

        try
        {
            // Invia request
            await SendRawAsync(request, cancellationToken).ConfigureAwait(false);

            // Attendi risposta con timeout
            var timeoutMs = request.TimeoutMs > 0 ? request.TimeoutMs : IpcConstants.RequestTimeoutMs;
            using var timeoutCts = new CancellationTokenSource(timeoutMs);
            using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeoutCts.Token);

            var registration = linkedCts.Token.Register(() =>
            {
                tcs.TrySetCanceled(linkedCts.Token);
            });

            try
            {
                return await tcs.Task.ConfigureAwait(false);
            }
            finally
            {
                await registration.DisposeAsync().ConfigureAwait(false);
            }
        }
        finally
        {
            _pendingRequests.TryRemove(request.CorrelationId, out _);
            _eventHandlers.TryRemove(request.CorrelationId, out _);
            _eventCounts.TryRemove(request.CorrelationId, out _);
        }
    }

    /// <summary>
    /// Invia richiesta di cancellazione per una operazione.
    /// Operazione idempotente - non fallisce se request già completata.
    /// </summary>
    /// <param name="correlationId">ID correlazione della request da cancellare.</param>
    /// <param name="cancellationToken">Token di cancellazione per l'invio.</param>
    public async Task SendCancelAsync(string correlationId, CancellationToken cancellationToken = default)
    {
        // Idempotente: non fare nulla se non connesso o già disposing
        if (!_isConnected || _isDisposing)
        {
            return;
        }

        try
        {
            var cancelRequest = new CancelRequest { CorrelationId = correlationId };
            await SendRawAsync(cancelRequest, cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            // Log ma non propagare - cancel è best-effort
            Debug.WriteLine($"[IpcClient] SendCancelAsync failed for {correlationId}: {ex.Message}");
        }
    }

    /// <summary>
    /// Invia heartbeat ping. Non attende response (gestita nel response loop).
    /// </summary>
    /// <param name="sequence">Numero di sequenza heartbeat.</param>
    /// <param name="cancellationToken">Token di cancellazione.</param>
    public async Task<HeartbeatPong?> SendHeartbeatAsync(long sequence, CancellationToken cancellationToken = default)
    {
        if (!_isConnected || _isDisposing)
        {
            return null;
        }

        try
        {
            var ping = new HeartbeatPing { Sequence = sequence };
            await SendRawAsync(ping, cancellationToken).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[IpcClient] SendHeartbeatAsync failed: {ex.Message}");
        }

        return null;
    }

    private async Task SendRawAsync<T>(T message, CancellationToken cancellationToken) where T : IpcMessage
    {
        await _sendLock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (_requestWriter == null || _isDisposing)
            {
                throw new InvalidOperationException("Not connected or disposing");
            }

            var json = JsonMessageSerializer.Serialize(message);

            // Verifica dimensione prima di inviare
            if (!IpcConstants.IsValidMessageSize(json.Length))
            {
                throw new InvalidOperationException($"Message exceeds maximum size ({json.Length} bytes)");
            }

            await _requestWriter.WriteLineAsync(json.AsMemory(), cancellationToken).ConfigureAwait(false);
            await _requestWriter.FlushAsync(cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            _sendLock.Release();
        }
    }

    private async Task<string?> ReadLineWithTimeoutAsync(StreamReader reader, CancellationToken cancellationToken)
    {
        try
        {
            return await reader.ReadLineAsync(cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            throw new TimeoutException("Read operation timed out");
        }
    }

    private async Task ResponseLoopAsync(CancellationToken cancellationToken)
    {
        try
        {
            while (!cancellationToken.IsCancellationRequested && _requestReader != null && !_isDisposing)
            {
                string? line;
                try
                {
                    line = await _requestReader.ReadLineAsync(cancellationToken).ConfigureAwait(false);
                }
                catch (IOException ex)
                {
                    Debug.WriteLine($"[IpcClient] ResponseLoop IOException: {ex.Message}");
                    break;
                }

                if (line == null)
                {
                    Debug.WriteLine("[IpcClient] ResponseLoop: pipe closed (null read)");
                    Console.WriteLine("[IpcClient] ResponseLoop: pipe closed (null read)");
                    break;
                }

                Console.WriteLine($"[IpcClient] Received response (length: {line.Length} bytes)");

                // Validazione dimensione messaggio
                if (!IpcConstants.IsValidMessageSize(line.Length))
                {
                    Debug.WriteLine($"[IpcClient] ResponseLoop: message too large ({line.Length} bytes), skipping");
                    Console.WriteLine($"[IpcClient] ResponseLoop: message too large ({line.Length} bytes), skipping");
                    continue;
                }

                IpcMessage? message;
                try
                {
                    message = JsonMessageSerializer.DeserializeMessage(line);
                    Console.WriteLine($"[IpcClient] Message deserialized: {message?.GetType().Name}");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[IpcClient] ResponseLoop: JSON parse error: {ex.Message}");
                    Console.WriteLine($"[IpcClient] ResponseLoop: JSON parse error: {ex.Message}");
                    continue;
                }

                if (message == null)
                {
                    Console.WriteLine("[IpcClient] Deserialized message is null, skipping");
                    continue;
                }

                switch (message)
                {
                    case ResponseEnvelope response:
                        Console.WriteLine($"[IpcClient] Processing ResponseEnvelope for correlation: {response.CorrelationId}");
                        Console.WriteLine($"[IpcClient] Response - Success: {response.Success}, WasCancelled: {response.WasCancelled}, HasPayload: {response.Payload != null}");
                        if (response.Error != null)
                        {
                            Console.WriteLine($"[IpcClient] Response has ERROR - Code: {response.Error.Code}, Message: {response.Error.Message}");
                        }
                        if (_pendingRequests.TryGetValue(response.CorrelationId, out var tcs))
                        {
                            Console.WriteLine($"[IpcClient] Found pending request, setting result");
                            tcs.TrySetResult(response);
                        }
                        else
                        {
                            Console.WriteLine($"[IpcClient] WARNING: No pending request found for correlation {response.CorrelationId}");
                        }
                        break;

                    case HeartbeatPong:
                        // Heartbeat ricevuto - worker è vivo
                        // Il supervisor gestisce il tracking
                        break;
                }
            }
        }
        catch (OperationCanceledException)
        {
            // Shutdown normale
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[IpcClient] ResponseLoop error: {ex.Message}");
        }

        // Solo se non stiamo già facendo dispose
        if (!_isDisposing)
        {
            await HandleDisconnectionAsync().ConfigureAwait(false);
        }
    }

    private async Task EventLoopAsync(CancellationToken cancellationToken)
    {
        try
        {
            while (!cancellationToken.IsCancellationRequested && _eventReader != null && !_isDisposing)
            {
                string? line;
                try
                {
                    line = await _eventReader.ReadLineAsync(cancellationToken).ConfigureAwait(false);
                }
                catch (IOException ex)
                {
                    Debug.WriteLine($"[IpcClient] EventLoop IOException: {ex.Message}");
                    break;
                }

                if (line == null)
                {
                    Debug.WriteLine("[IpcClient] EventLoop: pipe closed (null read)");
                    break;
                }

                // Validazione dimensione messaggio
                if (!IpcConstants.IsValidMessageSize(line.Length))
                {
                    Debug.WriteLine($"[IpcClient] EventLoop: message too large ({line.Length} bytes), skipping");
                    continue;
                }

                EventEnvelope? evt;
                try
                {
                    evt = JsonMessageSerializer.Deserialize<EventEnvelope>(line);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[IpcClient] EventLoop: JSON parse error: {ex.Message}");
                    continue;
                }

                if (evt == null)
                {
                    continue;
                }

                // Verifica limite eventi per request
                if (_eventCounts.TryGetValue(evt.CorrelationId, out var count))
                {
                    if (!IpcConstants.IsEventCountWithinLimit(count))
                    {
                        Debug.WriteLine($"[IpcClient] EventLoop: max events reached for {evt.CorrelationId}, dropping");
                        continue;
                    }
                    _eventCounts[evt.CorrelationId] = count + 1;
                }

                // Notifica handler specifico per correlazione
                if (_eventHandlers.TryGetValue(evt.CorrelationId, out var handler))
                {
                    try
                    {
                        handler(evt);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[IpcClient] EventHandler error: {ex.Message}");
                    }
                }

                // Notifica handler globale
                try
                {
                    EventReceived?.Invoke(this, evt);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[IpcClient] EventReceived handler error: {ex.Message}");
                }
            }
        }
        catch (OperationCanceledException)
        {
            // Shutdown normale
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[IpcClient] EventLoop error: {ex.Message}");
        }

        // Solo se non stiamo già facendo dispose
        if (!_isDisposing)
        {
            await HandleDisconnectionAsync().ConfigureAwait(false);
        }
    }

    private async Task HandleDisconnectionAsync()
    {
        lock (_stateLock)
        {
            if (!_isConnected)
            {
                return;
            }
            _isConnected = false;
        }

        // Cancella tutte le pending request con errore appropriato
        foreach (var kvp in _pendingRequests)
        {
            kvp.Value.TrySetException(new IOException("Connection to worker lost"));
        }
        _pendingRequests.Clear();
        _eventHandlers.Clear();
        _eventCounts.Clear();

        ConnectionStateChanged?.Invoke(this, WorkerConnectionState.Crashed);

        await DisposeInternalAsync().ConfigureAwait(false);
    }

    /// <summary>
    /// Disconnette dal worker in modo pulito.
    /// </summary>
    public async Task DisconnectAsync()
    {
        lock (_stateLock)
        {
            _isConnected = false;
        }

        _eventLoopCts?.Cancel();

        await DisposeInternalAsync().ConfigureAwait(false);

        ConnectionStateChanged?.Invoke(this, WorkerConnectionState.Stopped);
    }

    private async Task DisposeInternalAsync()
    {
        // Attendi che i loop terminino con timeout
        var loopTimeout = TimeSpan.FromMilliseconds(2000);

        try
        {
            if (_eventLoopTask != null)
            {
                await Task.WhenAny(_eventLoopTask, Task.Delay(loopTimeout)).ConfigureAwait(false);
            }
            if (_responseLoopTask != null)
            {
                await Task.WhenAny(_responseLoopTask, Task.Delay(loopTimeout)).ConfigureAwait(false);
            }
        }
        catch
        {
            // Ignora errori durante cleanup
        }

        _eventLoopCts?.Dispose();
        _eventLoopCts = null;

        // Dispose in ordine: prima readers/writers, poi pipes
        try { _requestReader?.Dispose(); } catch { }
        try { _requestWriter?.Dispose(); } catch { }
        try { _eventReader?.Dispose(); } catch { }
        try { _requestPipe?.Dispose(); } catch { }
        try { _eventPipe?.Dispose(); } catch { }

        _requestReader = null;
        _requestWriter = null;
        _eventReader = null;
        _requestPipe = null;
        _eventPipe = null;
    }

    private void ThrowIfDisposing()
    {
        if (_isDisposing)
        {
            throw new ObjectDisposedException(nameof(IpcClient));
        }
    }

    /// <summary>
    /// Rilascia le risorse asincronamente.
    /// </summary>
    public async ValueTask DisposeAsync()
    {
        if (_isDisposing)
        {
            return;
        }

        _isDisposing = true;

        await DisconnectAsync().ConfigureAwait(false);
        _sendLock.Dispose();
    }
}
