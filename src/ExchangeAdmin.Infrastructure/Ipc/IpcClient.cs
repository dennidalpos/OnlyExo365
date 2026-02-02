using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO.Pipes;
using System.Text;
using ExchangeAdmin.Contracts;
using ExchangeAdmin.Contracts.Messages;

namespace ExchangeAdmin.Infrastructure.Ipc;

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

    public event EventHandler<WorkerConnectionState>? ConnectionStateChanged;

    public event EventHandler<EventEnvelope>? EventReceived;

    public event EventHandler<HeartbeatPong>? HeartbeatReceived;

    public bool IsConnected => _isConnected;

    public IpcClient(string? pipeName = null, string? eventPipeName = null)
    {
        _pipeName = pipeName ?? IpcConstants.PipeName;
        _eventPipeName = eventPipeName ?? IpcConstants.EventPipeName;
    }

    public async Task<HandshakeResponse> ConnectAsync(CancellationToken cancellationToken = default)
    {
        ThrowIfDisposing();

        try
        {
            _requestPipe = new NamedPipeClientStream(".", _pipeName, PipeDirection.InOut, PipeOptions.Asynchronous);
            _eventPipe = new NamedPipeClientStream(".", _eventPipeName, PipeDirection.In, PipeOptions.Asynchronous);

            await Task.WhenAll(
                _requestPipe.ConnectAsync(IpcConstants.ConnectionTimeoutMs, cancellationToken),
                _eventPipe.ConnectAsync(IpcConstants.ConnectionTimeoutMs, cancellationToken)).ConfigureAwait(false);

            _requestReader = new StreamReader(_requestPipe, Encoding.UTF8, leaveOpen: true);
            _requestWriter = new StreamWriter(_requestPipe, Encoding.UTF8, leaveOpen: true);
            _eventReader = new StreamReader(_eventPipe, Encoding.UTF8, leaveOpen: true);

            var handshakeRequest = new HandshakeRequest();
            await SendRawAsync(handshakeRequest, cancellationToken).ConfigureAwait(false);

            var responseJson = await ReadLineWithTimeoutAsync(
                _requestReader,
                TimeSpan.FromMilliseconds(IpcConstants.HandshakeTimeoutMs),
                cancellationToken).ConfigureAwait(false);

            if (string.IsNullOrEmpty(responseJson))
            {
                throw new InvalidOperationException("Empty handshake response from worker");
            }

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

            if (!ContractVersion.IsCompatible(response.ContractsVersion))
            {
                throw new InvalidOperationException(
                    $"Contract version mismatch. Client: {ContractVersion.Version}, Worker: {response.ContractsVersion}");
            }

            lock (_stateLock)
            {
                _isConnected = true;
            }

            _eventLoopCts = new CancellationTokenSource();
            _eventLoopTask = Task.Run(() => EventLoopAsync(_eventLoopCts.Token), CancellationToken.None);
            _responseLoopTask = Task.Run(() => ResponseLoopAsync(_eventLoopCts.Token), CancellationToken.None);

            ConnectionStateChanged?.Invoke(this, WorkerConnectionState.Connected);

            return response;
        }
        catch (OperationCanceledException)
        {
            await DisposeInternalAsync().ConfigureAwait(false);
            throw;
        }
        catch (Exception)
        {
            await DisposeInternalAsync().ConfigureAwait(false);
            throw;
        }
    }

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

        if (eventHandler != null)
        {
            _eventHandlers[request.CorrelationId] = eventHandler;
            _eventCounts[request.CorrelationId] = 0;
        }

        _pendingRequests[request.CorrelationId] = tcs;

        try
        {
            await SendRawAsync(request, cancellationToken).ConfigureAwait(false);

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

    public async Task SendCancelAsync(string correlationId, CancellationToken cancellationToken = default)
    {
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
            Debug.WriteLine($"[IpcClient] SendCancelAsync failed for {correlationId}: {ex.Message}");
        }
    }

    public async Task SendHeartbeatAsync(long sequence, CancellationToken cancellationToken = default)
    {
        if (!_isConnected || _isDisposing)
        {
            return;
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

    private async Task<string?> ReadLineWithTimeoutAsync(
        StreamReader reader,
        TimeSpan timeout,
        CancellationToken cancellationToken)
    {
        using var timeoutCts = new CancellationTokenSource(timeout);
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeoutCts.Token);

        try
        {
            return await reader.ReadLineAsync(linkedCts.Token).ConfigureAwait(false);
        }
        catch (OperationCanceledException) when (timeoutCts.IsCancellationRequested && !cancellationToken.IsCancellationRequested)
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
                    break;
                }

                if (!IpcConstants.IsValidMessageSize(line.Length))
                {
                    Debug.WriteLine($"[IpcClient] ResponseLoop: message too large ({line.Length} bytes), skipping");
                    continue;
                }

                IpcMessage? message;
                try
                {
                    message = JsonMessageSerializer.DeserializeMessage(line);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[IpcClient] ResponseLoop: JSON parse error: {ex.Message}");
                    continue;
                }

                if (message == null)
                {
                    continue;
                }

                switch (message)
                {
                    case ResponseEnvelope response:
                        if (_pendingRequests.TryGetValue(response.CorrelationId, out var tcs))
                        {
                            tcs.TrySetResult(response);
                        }
                        break;

                    case HeartbeatPong pong:
                        HeartbeatReceived?.Invoke(this, pong);
                        break;
                }
            }
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[IpcClient] ResponseLoop error: {ex.Message}");
        }

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

                if (_eventCounts.TryGetValue(evt.CorrelationId, out var count))
                {
                    if (!IpcConstants.IsEventCountWithinLimit(count))
                    {
                        Debug.WriteLine($"[IpcClient] EventLoop: max events reached for {evt.CorrelationId}, dropping");
                        continue;
                    }
                    _eventCounts[evt.CorrelationId] = count + 1;
                }

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
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[IpcClient] EventLoop error: {ex.Message}");
        }

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
        }

        _eventLoopCts?.Dispose();
        _eventLoopCts = null;

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
