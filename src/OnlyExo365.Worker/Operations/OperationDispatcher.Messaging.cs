using OnlyExo365.Contracts.Messages;

namespace OnlyExo365.Worker.Operations;

public partial class OperationDispatcher
{
    private LogLevel ParseLogLevel(string level)
    {
        return _eventPublisher.ParseLogLevel(level);
    }

    private Task SendLogAsync(string correlationId, LogLevel level, string message)
    {
        return _eventPublisher.SendLogAsync(correlationId, level, message);
    }

    private Task SendProgressAsync(string correlationId, int percentComplete, string? statusMessage, int? currentItem = null, int? totalItems = null)
    {
        return _eventPublisher.SendProgressAsync(correlationId, percentComplete, statusMessage, currentItem, totalItems);
    }

    private Task SendPartialOutputAsync<T>(string correlationId, T data, int itemIndex)
    {
        return _eventPublisher.SendPartialOutputAsync(correlationId, data, itemIndex);
    }

    private ResponseEnvelope CreateSuccessResponse<T>(string correlationId, T payload)
    {
        return _responseFactory.CreateSuccess(correlationId, payload);
    }

    private ResponseEnvelope CreateErrorResponse(string correlationId, ErrorCode code, string message, bool isTransient = false, int? retryAfterSeconds = null)
    {
        return _responseFactory.CreateError(correlationId, code, message, isTransient, retryAfterSeconds);
    }

    private ResponseEnvelope CreateCancelledResponse(string correlationId)
    {
        return _responseFactory.CreateCancelled(correlationId);
    }
}

