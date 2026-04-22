using System.Text.Json;
using OnlyExo365.Contracts;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Errors;
using OnlyExo365.Contracts.Results;

namespace OnlyExo365.Shell.Ipc;

internal sealed class WorkerClientRuntime
{
    private readonly WorkerSupervisor _supervisor;
    private readonly WorkerOperationResiliencePipeline _resiliencePipeline;

    public WorkerClientRuntime(WorkerSupervisor supervisor, WorkerOperationResiliencePipeline? resiliencePipeline = null)
    {
        _supervisor = supervisor;
        _resiliencePipeline = resiliencePipeline ?? new WorkerOperationResiliencePipeline();
    }

    public WorkerConnectionState State => _supervisor.State;
    public WorkerStatus Status => _supervisor.GetStatus();

    public void UpdateWorkerConsoleVisibility(bool isVisible)
        => _supervisor.SetConsoleVisibility(isVisible);

    public async Task<Result<TResponse>> ExecuteOperationAsync<TResponse>(
        OperationType operation,
        object? payload,
        Action<EventEnvelope>? eventHandler,
        CancellationToken cancellationToken)
    {
        return await _resiliencePipeline.ExecuteAsync(
            operation,
            ct => SendRequestForOperationAsync<TResponse>(operation, payload, eventHandler, ct),
            cancellationToken);
    }

    public async Task<Result> ExecuteCommandAsync(
        OperationType operation,
        object? payload,
        Action<EventEnvelope>? eventHandler,
        CancellationToken cancellationToken)
    {
        return await _resiliencePipeline.ExecuteAsync(
            operation,
            ct => SendRequestForCommandAsync(operation, payload, eventHandler, ct),
            cancellationToken);
    }

    public async Task CancelOperationAsync(string correlationId)
    {
        if (_supervisor.State != WorkerConnectionState.Connected)
            return;

        await _supervisor.IpcClient.SendCancelAsync(correlationId);
    }

    private async Task<ResponseEnvelope> SendRequestInternalAsync(
        OperationType operation,
        object? payload,
        Action<EventEnvelope>? eventHandler,
        CancellationToken cancellationToken)
    {
        if (_supervisor.State != WorkerConnectionState.Connected)
        {
            return new ResponseEnvelope
            {
                Success = false,
                Error = new NormalizedErrorDto
                {
                    Code = ErrorCode.WorkerNotRunning,
                    Message = "Worker is not running",
                    IsTransient = false
                }
            };
        }

        using var preparedPayload = PreparedIpcPayload.Create(payload);
        var request = CreateRequestEnvelope(operation, preparedPayload.Payload);

        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var correlationId = request.CorrelationId;

        var registration = cancellationToken.Register(async () =>
        {
            await _supervisor.IpcClient.SendCancelAsync(correlationId);
        });

        try
        {
            return await _supervisor.IpcClient.SendRequestAsync(request, eventHandler, linkedCts.Token);
        }
        catch (OperationCanceledException)
        {
            return new ResponseEnvelope
            {
                CorrelationId = correlationId,
                Success = false,
                WasCancelled = true
            };
        }
        finally
        {
            await registration.DisposeAsync();
        }
    }

    internal static RequestEnvelope CreateRequestEnvelope(OperationType operation, object? payload)
    {
        return new RequestEnvelope
        {
            Operation = operation,
            Payload = payload != null ? JsonMessageSerializer.ToJsonElement(payload) : null,
            TimeoutMs = IpcConstants.RequestTimeoutMs
        };
    }

    private async Task<Result<TResponse>> SendRequestForOperationAsync<TResponse>(
        OperationType operation,
        object? payload,
        Action<EventEnvelope>? eventHandler,
        CancellationToken cancellationToken)
    {
        try
        {
            var response = await SendRequestInternalAsync(operation, payload, eventHandler, cancellationToken);

            if (response.WasCancelled)
                return Result<TResponse>.Cancelled(response.CorrelationId);

            if (!response.Success)
                return Result<TResponse>.Failure(NormalizedError.FromDto(response.Error!), response.CorrelationId);

            if (response.Payload == null)
                return Result<TResponse>.Success(default!, response.CorrelationId);

            var result = JsonMessageSerializer.ExtractPayload<TResponse>(response.Payload);
            return Result<TResponse>.Success(result!, response.CorrelationId);
        }
        catch (Exception ex)
        {
            return Result<TResponse>.FromException(ex);
        }
    }

    private async Task<Result> SendRequestForCommandAsync(
        OperationType operation,
        object? payload,
        Action<EventEnvelope>? eventHandler,
        CancellationToken cancellationToken)
    {
        try
        {
            var response = await SendRequestInternalAsync(operation, payload, eventHandler, cancellationToken);

            if (response.WasCancelled)
                return Result.Cancelled(response.CorrelationId);

            if (!response.Success)
                return Result.Failure(NormalizedError.FromDto(response.Error!), response.CorrelationId);

            return Result.Success(response.CorrelationId);
        }
        catch (Exception ex)
        {
            return Result.FromException(ex);
        }
    }

}

