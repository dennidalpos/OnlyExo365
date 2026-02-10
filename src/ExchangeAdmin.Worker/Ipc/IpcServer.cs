using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO.Pipes;
using System.Text;
using ExchangeAdmin.Contracts;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Worker.Operations;
using ExchangeAdmin.Worker.PowerShell;

namespace ExchangeAdmin.Worker.Ipc;

             
                                         
                                                                                
              
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

                 
                                 
                  
                                                                                
    public IpcServer(PowerShellEngine psEngine)
    {
        _psEngine = psEngine;
        _dispatcher = new OperationDispatcher(psEngine, SendEventAsync);
    }

                 
                                                         
                  
    public async Task StartAsync()
    {
        _serverCts = new CancellationTokenSource();

                                                    
        _requestPipe = new NamedPipeServerStream(
            IpcConstants.PipeName,
            PipeDirection.InOut,
            1,
            PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous,
            IpcConstants.PipeBufferSize,
            IpcConstants.PipeBufferSize);

                           
        _eventPipe = new NamedPipeServerStream(
            IpcConstants.EventPipeName,
            PipeDirection.Out,
            1,
            PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous,
            IpcConstants.PipeBufferSize,
            0);

        ConsoleLogger.Info("IPC", $"Waiting for client connection on {IpcConstants.PipeName}...");

        await Task.WhenAll(
            _requestPipe.WaitForConnectionAsync(_serverCts.Token),
            _eventPipe.WaitForConnectionAsync(_serverCts.Token)).ConfigureAwait(false);

        ConsoleLogger.Success("IPC", "Client connected");

        _requestReader = new StreamReader(_requestPipe, Encoding.UTF8, leaveOpen: true);
        _requestWriter = new StreamWriter(_requestPipe, Encoding.UTF8, leaveOpen: true);
        _eventWriter = new StreamWriter(_eventPipe, Encoding.UTF8, leaveOpen: true);

        ConsoleLogger.Debug("IPC", "Readers/writers created successfully");

        _isRunning = true;

                               
        _requestLoopTask = RequestLoopAsync(_serverCts.Token);
    }

                 
                                           
                  
    public async Task StopAsync()
    {
        _isRunning = false;
        _serverCts?.Cancel();

                                              
        foreach (var kvp in _activeOperations)
        {
            try
            {
                kvp.Value.Cancel();
            }
            catch (ObjectDisposedException)
            {
                                
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
                    ConsoleLogger.Error("IPC", $"RequestLoop IOException: {ex.Message}");
                    break;
                }

                if (line == null)
                {
                    ConsoleLogger.Warning("IPC", "Client disconnected (null read)");
                    break;
                }

                if (!IpcConstants.IsValidMessageSize(line.Length))
                {
                    ConsoleLogger.Warning("IPC", $"Message too large ({line.Length} bytes), rejecting");
                    continue;
                }

                IpcMessage? message;
                try
                {
                    message = JsonMessageSerializer.DeserializeMessage(line);
                }
                catch (Exception ex)
                {
                    ConsoleLogger.Error("IPC", $"JSON parse error: {ex.Message}");
                    continue;
                }

                if (message == null)
                {
                    ConsoleLogger.Warning("IPC", "Invalid message received (unknown type)");
                    continue;
                }

                                                                          
                _ = HandleMessageAsync(message, cancellationToken);
            }
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            ConsoleLogger.Error("IPC", $"Request loop fatal error: {ex.Message}");
        }
        finally
        {
            ConsoleLogger.Info("IPC", "Request loop terminated");
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
                    ConsoleLogger.Warning("IPC", $"Unknown message type: {message.Type}");
                    break;
            }
        }
        catch (Exception ex)
        {
            ConsoleLogger.Error("IPC", $"Error handling message {message.Type}: {ex.Message}");
        }
    }

    private async Task HandleHandshakeAsync(HandshakeRequest request)
    {
        ConsoleLogger.Info("IPC", $"Handshake request from client {request.ClientId}");
        ConsoleLogger.Debug("IPC", $"Client contracts version: {request.ContractsVersion}");

        _clientId = request.ClientId;

        var isCompatible = ContractVersion.IsCompatible(request.ContractsVersion);

        var response = new HandshakeResponse
        {
            Success = isCompatible,
            ContractsVersion = ContractVersion.Version,
            WorkerVersion = "1.0.1",
            IsModuleAvailable = _psEngine.IsModuleAvailable,
            PowerShellVersion = _psEngine.PowerShellVersion,
            ErrorMessage = isCompatible ? null : $"Incompatible contracts version. Worker: {ContractVersion.Version}, Client: {request.ContractsVersion}"
        };

        ConsoleLogger.Verbose("IPC", "Sending handshake response...");
        await SendResponseRawAsync(response).ConfigureAwait(false);

        if (isCompatible)
        {
            ConsoleLogger.Success("IPC", "Handshake completed successfully");
        }
        else
        {
            ConsoleLogger.Error("IPC", $"Handshake failed - incompatible version");
        }
    }

    private async Task HandleRequestAsync(RequestEnvelope request, CancellationToken serverCancellation)
    {
        ConsoleLogger.Info("IPC", $"Request: {request.Operation} (correlation: {request.CorrelationId})");

                                                          
        _eventCounts[request.CorrelationId] = 0;

                                                       
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
            ConsoleLogger.Error("IPC", $"Operation error: {ex.Message}");

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
                                            
                }

                operationCts = null;
            }
        }
    }

    private void HandleCancel(CancelRequest cancel)
    {
        ConsoleLogger.Warning("IPC", $"Cancel request for correlation: {cancel.CorrelationId}");

        if (_activeOperations.TryGetValue(cancel.CorrelationId, out var cts))
        {
            try
            {
                if (!cts.IsCancellationRequested)
                {
                    cts.Cancel();
                    ConsoleLogger.Warning("IPC", $"Cancellation signaled for: {cancel.CorrelationId}");
                }
                else
                {
                    ConsoleLogger.Debug("IPC", $"Already cancelled: {cancel.CorrelationId}");
                }
            }
            catch (ObjectDisposedException)
            {
                ConsoleLogger.Debug("IPC", $"Operation already completed: {cancel.CorrelationId}");
            }
        }
        else
        {
            ConsoleLogger.Debug("IPC", $"Cancel ignored - operation not found: {cancel.CorrelationId}");
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
            ConsoleLogger.Warning("IPC", "Cannot send response - writer is null or disposing");
            return;
        }

        try
        {
            var json = JsonMessageSerializer.Serialize(message);

            if (!IpcConstants.IsValidMessageSize(json.Length))
            {
                ConsoleLogger.Error("IPC", $"Response too large ({json.Length} bytes), dropping");
                return;
            }

            ConsoleLogger.Verbose("IPC", $"Sending response ({json.Length} bytes, type: {typeof(T).Name})");
            await _requestWriter.WriteLineAsync(json).ConfigureAwait(false);
            await _requestWriter.FlushAsync().ConfigureAwait(false);
            ConsoleLogger.Debug("IPC", "Response sent successfully");
        }
        catch (IOException ex)
        {
            ConsoleLogger.Error("IPC", $"Failed to send response: {ex.Message}");
        }
        catch (ObjectDisposedException)
        {
                               
        }
    }

    private async Task SendEventAsync(EventEnvelope evt)
    {
        if (_eventWriter == null || _isDisposing)
        {
            return;
        }

                                             
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

                                                   
            if (!IpcConstants.IsValidMessageSize(json.Length))
            {
                ConsoleLogger.Error("IPC", $"Event too large ({json.Length} bytes), dropping");
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

                 
                            
                  
    public void Dispose()
    {
        if (_isDisposing)
        {
            return;
        }

        _isDisposing = true;

        _serverCts?.Cancel();

                                              
        foreach (var kvp in _activeOperations)
        {
            try
            {
                kvp.Value.Cancel();
                kvp.Value.Dispose();
            }
            catch
            {
                         
            }
        }
        _activeOperations.Clear();

        _serverCts?.Dispose();
        _eventWriteLock.Dispose();
        Cleanup();
    }
}
