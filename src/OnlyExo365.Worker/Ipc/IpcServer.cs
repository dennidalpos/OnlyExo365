using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO.Pipes;
using System.Text;
using System.Threading.Channels;
using OnlyExo365.Contracts;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Worker.Operations;
using OnlyExo365.Worker.PowerShell;

namespace OnlyExo365.Worker.Ipc;

             
                                         
                                                                                
              
public sealed class IpcServer : IDisposable
{
    private readonly PowerShellEngine _psEngine;
    private readonly OperationDispatcher _dispatcher;
    private readonly WorkerConsoleController _consoleController;
    private readonly IpcSessionContext _sessionContext;
    private readonly string _expectedSessionToken;
    private readonly DateTime _startTime = DateTime.UtcNow;

    private NamedPipeServerStream? _requestPipe;
    private NamedPipeServerStream? _eventPipe;
    private StreamReader? _requestReader;
    private StreamWriter? _requestWriter;
    private StreamWriter? _eventWriter;

    private CancellationTokenSource? _serverCts;
    private Task? _requestLoopTask;
    private Task? _requestProcessingTask;
    private Channel<RequestEnvelope>? _pendingRequests;

    private readonly ConcurrentDictionary<string, CancellationTokenSource> _activeOperations = new();
    private readonly ConcurrentDictionary<string, Task> _backgroundRequests = new();
    private readonly ConcurrentDictionary<string, int> _eventCounts = new();
    private readonly SemaphoreSlim _requestWriteLock = new(1, 1);
    private readonly SemaphoreSlim _eventWriteLock = new(1, 1);

    private volatile bool _isRunning;
    private volatile bool _isDisposing;
    private volatile bool _handshakeCompleted;
    private string? _clientId;

                 
                                 
                  
                                                                                
    public IpcServer(PowerShellEngine psEngine, IpcSessionContext sessionContext, string expectedSessionToken)
    {
        _psEngine = psEngine;
        _sessionContext = sessionContext ?? throw new ArgumentNullException(nameof(sessionContext));
        _expectedSessionToken = string.IsNullOrWhiteSpace(expectedSessionToken)
            ? throw new ArgumentException("Expected IPC session token is required.", nameof(expectedSessionToken))
            : expectedSessionToken;
        _consoleController = new WorkerConsoleController();
        _dispatcher = new OperationDispatcher(psEngine, SendEventAsync, _consoleController);
    }

                 
                                                         
                  
    public async Task StartAsync()
    {
        _serverCts = new CancellationTokenSource();

                                                    
        _requestPipe = new NamedPipeServerStream(
            _sessionContext.RequestPipeName,
            PipeDirection.InOut,
            1,
            PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly,
            IpcConstants.PipeBufferSize,
            IpcConstants.PipeBufferSize);

                           
        _eventPipe = new NamedPipeServerStream(
            _sessionContext.EventPipeName,
            PipeDirection.Out,
            1,
            PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly,
            IpcConstants.PipeBufferSize,
            0);

        ConsoleLogger.Info("IPC", $"Waiting for client connection on {_sessionContext.RequestPipeName}...");

        await Task.WhenAll(
            _requestPipe.WaitForConnectionAsync(_serverCts.Token),
            _eventPipe.WaitForConnectionAsync(_serverCts.Token)).ConfigureAwait(false);

        ConsoleLogger.Success("IPC", "Client connected");

        _requestReader = new StreamReader(_requestPipe, Encoding.UTF8, leaveOpen: true);
        _requestWriter = new StreamWriter(_requestPipe, Encoding.UTF8, leaveOpen: true);
        _eventWriter = new StreamWriter(_eventPipe, Encoding.UTF8, leaveOpen: true);

        ConsoleLogger.Debug("IPC", "Readers/writers created successfully");

        _pendingRequests = CreatePendingRequestChannel();
        _isRunning = true;

                               
        _requestLoopTask = RequestLoopAsync(_serverCts.Token);
        _requestProcessingTask = ProcessRequestQueueAsync(_serverCts.Token);
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
        _handshakeCompleted = false;

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

        if (_requestProcessingTask != null)
        {
            try
            {
                await Task.WhenAny(_requestProcessingTask, Task.Delay(5000)).ConfigureAwait(false);
            }
            catch
            {
            }
        }

        if (_backgroundRequests.Count > 0)
        {
            try
            {
                await Task.WhenAny(Task.WhenAll(_backgroundRequests.Values), Task.Delay(5000)).ConfigureAwait(false);
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

                                                                          
                await HandleMessageAsync(message, cancellationToken).ConfigureAwait(false);
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
                    if (!EnsureHandshakeCompleted())
                    {
                        return;
                    }
                    await EnqueueRequestAsync(request, cancellationToken).ConfigureAwait(false);
                    break;

                case CancelRequest cancel:
                    if (!EnsureHandshakeCompleted())
                    {
                        return;
                    }
                    HandleCancel(cancel);
                    break;

                case HeartbeatPing ping:
                    if (!EnsureHandshakeCompleted())
                    {
                        return;
                    }
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

        var isCompatible = ContractVersion.IsCompatible(request.ContractsVersion);
        var sessionValidated = isCompatible && ValidateHandshake(request);

        var response = new HandshakeResponse
        {
            Success = isCompatible && sessionValidated,
            ContractsVersion = ContractVersion.Version,
            WorkerVersion = ProductInfo.Version,
            IsModuleAvailable = _psEngine.IsModuleAvailable,
            PowerShellVersion = _psEngine.PowerShellVersion,
            SessionValidated = sessionValidated,
            ErrorMessage = !isCompatible
                ? $"Incompatible contracts version. Worker: {ContractVersion.Version}, Client: {request.ContractsVersion}"
                : sessionValidated
                    ? null
                    : "IPC session validation failed."
        };

        ConsoleLogger.Verbose("IPC", "Sending handshake response...");
        await SendResponseRawAsync(response).ConfigureAwait(false);

        if (response.Success)
        {
            _clientId = request.ClientId;
            _handshakeCompleted = true;
            ConsoleLogger.Success("IPC", "Handshake completed successfully");
        }
        else
        {
            ConsoleLogger.Error("IPC", $"Handshake failed - {response.ErrorMessage}");
        }
    }

    private bool ValidateHandshake(HandshakeRequest request)
    {
        if (!string.Equals(request.SessionToken, _expectedSessionToken, StringComparison.Ordinal))
        {
            ConsoleLogger.Warning("IPC", "Handshake rejected: invalid bootstrap token.");
            return false;
        }

        var requestContext = new IpcSessionContext
        {
            SessionId = request.SessionId,
            UserScope = request.UserScope
        };

        if (!_sessionContext.Matches(requestContext))
        {
            ConsoleLogger.Warning("IPC", "Handshake rejected: session binding mismatch.");
            return false;
        }

        return true;
    }

    private bool EnsureHandshakeCompleted()
    {
        if (_handshakeCompleted)
        {
            return true;
        }

        ConsoleLogger.Warning("IPC", "Request rejected before successful handshake.");
        return false;
    }

    private async Task EnqueueRequestAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        if (_pendingRequests is null)
        {
            throw new InvalidOperationException("Request queue is not initialized.");
        }

        var queueDepth = _pendingRequests.Reader.Count;
        if (queueDepth >= IpcConstants.MaxPendingRequests)
        {
            ConsoleLogger.Warning("IPC", $"Request queue full ({queueDepth}/{IpcConstants.MaxPendingRequests}), applying backpressure.");
        }

        await _pendingRequests.Writer.WriteAsync(request, cancellationToken).ConfigureAwait(false);
        ConsoleLogger.Debug("IPC", $"Request queued: {request.Operation} (pending: {_pendingRequests.Reader.Count})");
    }

    private async Task ProcessRequestQueueAsync(CancellationToken cancellationToken)
    {
        if (_pendingRequests is null)
        {
            return;
        }

        try
        {
            await foreach (var request in _pendingRequests.Reader.ReadAllAsync(cancellationToken).ConfigureAwait(false))
            {
                if (CanProcessConcurrently(request.Operation))
                {
                    StartConcurrentRequest(request, cancellationToken);
                    continue;
                }

                await HandleRequestAsync(request, cancellationToken).ConfigureAwait(false);
            }
        }
        catch (OperationCanceledException)
        {
        }
        catch (ChannelClosedException)
        {
        }
        catch (Exception ex)
        {
            ConsoleLogger.Error("IPC", $"Request processing loop fatal error: {ex.Message}");
        }
        finally
        {
            ConsoleLogger.Info("IPC", "Request processing loop terminated");
        }
    }

    internal static Channel<RequestEnvelope> CreatePendingRequestChannel(int capacity = IpcConstants.MaxPendingRequests)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(capacity);

        return Channel.CreateBounded<RequestEnvelope>(new BoundedChannelOptions(capacity)
        {
            FullMode = BoundedChannelFullMode.Wait,
            SingleReader = true,
            SingleWriter = true
        });
    }

    internal static bool CanProcessConcurrently(OperationType operation)
        => operation == OperationType.SearchUnifiedAuditLog;

    private void StartConcurrentRequest(RequestEnvelope request, CancellationToken cancellationToken)
    {
        ConsoleLogger.Debug("IPC", $"Dispatching concurrent request: {request.Operation} ({request.CorrelationId})");

        var task = Task.Run(() => HandleRequestAsync(request, cancellationToken), CancellationToken.None);
        _backgroundRequests[request.CorrelationId] = task;
        _ = ObserveConcurrentRequestAsync(request.CorrelationId, task);
    }

    private async Task ObserveConcurrentRequestAsync(string correlationId, Task task)
    {
        try
        {
            await task.ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            ConsoleLogger.Error("IPC", $"Concurrent request {correlationId} failed: {ex.Message}");
        }
        finally
        {
            _backgroundRequests.TryRemove(correlationId, out _);
        }
    }

    private async Task HandleRequestAsync(RequestEnvelope request, CancellationToken serverCancellation)
    {
        using var correlationScope = ConsoleLogger.BeginCorrelationScope(request.CorrelationId);
        ConsoleLogger.Info("IPC", $"Request: {request.Operation} (correlation: {request.CorrelationId})");

                                                          
        _eventCounts[request.CorrelationId] = 0;

                                                       
        CancellationTokenSource? operationCts = null;
        try
        {
            operationCts = CancellationTokenSource.CreateLinkedTokenSource(serverCancellation);
            _activeOperations[request.CorrelationId] = operationCts;

            var response = await _dispatcher.DispatchAsync(request, operationCts.Token).ConfigureAwait(false);
            await SendOperationResponseAsync(request.CorrelationId, response).ConfigureAwait(false);
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
            await SendOperationResponseAsync(request.CorrelationId, response).ConfigureAwait(false);
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
            await SendOperationResponseAsync(request.CorrelationId, response).ConfigureAwait(false);
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

    private Task SendOperationResponseAsync(string correlationId, ResponseEnvelope response)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(correlationId);
        ArgumentNullException.ThrowIfNull(response);

        response.CorrelationId = correlationId;
        return SendResponseRawAsync(response, correlationId);
    }

    private async Task SendResponseRawAsync<T>(T message, string? operationCorrelationId = null) where T : IpcMessage
    {
        if (_requestWriter == null || _isDisposing)
        {
            ConsoleLogger.Warning("IPC", "Cannot send response - writer is null or disposing");
            return;
        }

        try
        {
            var json = SerializeResponseForTransport(message, operationCorrelationId);

            ConsoleLogger.Verbose("IPC", $"Sending response ({json.Length} bytes, type: {typeof(T).Name})");
            await WriteSerializedMessageAsync(_requestWriter, _requestWriteLock, json).ConfigureAwait(false);
            ConsoleLogger.Debug("IPC", "Response sent successfully");
        }
        catch (ObjectDisposedException)
        {
                               
        }
        catch (InvalidOperationException ex)
        {
            ConsoleLogger.Error("IPC", $"Failed to prepare response: {ex.Message}");
        }
        catch (IOException ex)
        {
            ConsoleLogger.Error("IPC", $"Failed to send response: {ex.Message}");
        }
    }

    internal static string SerializeResponseForTransport<T>(T message, string? operationCorrelationId = null) where T : IpcMessage
    {
        var json = JsonMessageSerializer.Serialize(message);
        if (IpcConstants.IsValidMessageSize(json.Length))
        {
            return json;
        }

        var correlationId = message switch
        {
            ResponseEnvelope response when !string.IsNullOrWhiteSpace(response.CorrelationId) => response.CorrelationId,
            _ when !string.IsNullOrWhiteSpace(operationCorrelationId) => operationCorrelationId,
            _ => null
        };

        if (string.IsNullOrWhiteSpace(correlationId))
        {
            throw new InvalidOperationException($"Response too large ({json.Length} bytes) and no correlation id was available for fallback.");
        }

        var oversizeError = CreateOversizeResponse(correlationId, json.Length);
        var oversizeJson = JsonMessageSerializer.Serialize(oversizeError);

        if (!IpcConstants.IsValidMessageSize(oversizeJson.Length))
        {
            throw new InvalidOperationException($"Oversize fallback response also exceeded the maximum size ({oversizeJson.Length} bytes).");
        }

        return oversizeJson;
    }

    internal static ResponseEnvelope CreateOversizeResponse(string correlationId, int attemptedSizeBytes)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(correlationId);

        return new ResponseEnvelope
        {
            CorrelationId = correlationId,
            Success = false,
            Error = new NormalizedErrorDto
            {
                Code = ErrorCode.MessageTooLarge,
                Message = $"Worker response exceeded the maximum IPC message size ({attemptedSizeBytes} bytes > {IpcConstants.MaxMessageSizeBytes} bytes).",
                IsTransient = false
            }
        };
    }

    internal static async Task WriteSerializedMessageAsync(TextWriter writer, SemaphoreSlim writeLock, string message, CancellationToken cancellationToken = default)
    {
        await writeLock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            await writer.WriteLineAsync(message.AsMemory(), cancellationToken).ConfigureAwait(false);
            await writer.FlushAsync(cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            writeLock.Release();
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
        _pendingRequests?.Writer.TryComplete();
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
        _pendingRequests = null;
        _clientId = null;
        _backgroundRequests.Clear();
        _handshakeCompleted = false;
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
        _backgroundRequests.Clear();

        _serverCts?.Dispose();
        _requestWriteLock.Dispose();
        _eventWriteLock.Dispose();
        Cleanup();
    }
}

