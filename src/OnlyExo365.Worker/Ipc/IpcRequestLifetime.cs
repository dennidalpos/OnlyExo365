using OnlyExo365.Contracts;
using OnlyExo365.Contracts.Messages;

namespace OnlyExo365.Worker.Ipc;

internal sealed class IpcRequestLifetime : IDisposable
{
    private readonly CancellationTokenSource _cancellation = new();
    private readonly CancellationTokenSource _deadline;
    private readonly CancellationTokenSource _linked;

    public IpcRequestLifetime(int timeoutMs, CancellationToken serverCancellation)
    {
        _deadline = new CancellationTokenSource(timeoutMs > 0 ? timeoutMs : IpcConstants.RequestTimeoutMs);
        _linked = CancellationTokenSource.CreateLinkedTokenSource(serverCancellation, _cancellation.Token, _deadline.Token);
    }

    public CancellationToken Token => _linked.Token;
    public bool IsCancellationRequested => _linked.IsCancellationRequested;
    public bool IsTimedOut => _deadline.IsCancellationRequested && !_cancellation.IsCancellationRequested;
    public void Cancel()
    {
        if (!_deadline.IsCancellationRequested)
        {
            _cancellation.Cancel();
        }
    }

    public ResponseEnvelope CreateCancellationResponse(string correlationId) => new()
    {
        CorrelationId = correlationId,
        Success = false,
        WasCancelled = !IsTimedOut,
        Error = IsTimedOut ? new NormalizedErrorDto
        {
            Code = ErrorCode.Timeout,
            Message = "The worker request deadline expired. Any changes already submitted may have completed.",
            IsTransient = true
        } : null
    };

    public void Dispose()
    {
        _linked.Dispose();
        _deadline.Dispose();
        _cancellation.Dispose();
    }
}
