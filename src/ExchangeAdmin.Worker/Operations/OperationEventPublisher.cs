using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;

namespace ExchangeAdmin.Worker.Operations;

internal sealed class OperationEventPublisher(Func<EventEnvelope, Task> sendEvent)
{
    public LogLevel ParseLogLevel(string level)
    {
        return level.ToLowerInvariant() switch
        {
            "verbose" => LogLevel.Verbose,
            "debug" => LogLevel.Debug,
            "information" or "info" => LogLevel.Information,
            "warning" or "warn" => LogLevel.Warning,
            "error" => LogLevel.Error,
            _ => LogLevel.Information
        };
    }

    public Task SendLogAsync(string correlationId, LogLevel level, string message)
    {
        var payload = new LogEventPayload
        {
            Level = level,
            Message = message,
            Source = "Worker"
        };

        return SendAsync(correlationId, EventType.Log, payload);
    }

    public Task SendProgressAsync(string correlationId, int percentComplete, string? statusMessage, int? currentItem = null, int? totalItems = null)
    {
        var payload = new ProgressEventPayload
        {
            PercentComplete = percentComplete,
            StatusMessage = statusMessage,
            CurrentItem = currentItem,
            TotalItems = totalItems
        };

        return SendAsync(correlationId, EventType.Progress, payload);
    }

    public Task SendPartialOutputAsync<T>(string correlationId, T data, int itemIndex)
    {
        var payload = new PartialOutputPayload
        {
            Data = JsonMessageSerializer.ToJsonElement(data),
            ItemIndex = itemIndex
        };

        return SendAsync(correlationId, EventType.PartialOutput, payload);
    }

    private Task SendAsync<TPayload>(string correlationId, EventType eventType, TPayload payload)
    {
        var evt = new EventEnvelope
        {
            CorrelationId = correlationId,
            EventType = eventType,
            Payload = JsonMessageSerializer.ToJsonElement(payload)
        };

        return sendEvent(evt);
    }
}
