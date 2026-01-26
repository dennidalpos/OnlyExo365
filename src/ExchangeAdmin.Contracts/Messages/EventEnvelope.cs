using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Messages;

             
                                             
              
public class EventEnvelope : IpcMessage
{
    public EventEnvelope()
    {
        Type = MessageType.Event;
    }

    [JsonPropertyName("correlationId")]
    public string CorrelationId { get; set; } = string.Empty;

    [JsonPropertyName("eventType")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public EventType EventType { get; set; }

    [JsonPropertyName("payload")]
    public JsonElement? Payload { get; set; }
}

             
                              
              
public class LogEventPayload
{
    [JsonPropertyName("level")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public LogLevel Level { get; set; }

    [JsonPropertyName("message")]
    public string Message { get; set; } = string.Empty;

    [JsonPropertyName("source")]
    public string? Source { get; set; }
}

             
                                   
              
public class ProgressEventPayload
{
    [JsonPropertyName("percentComplete")]
    public int PercentComplete { get; set; }

    [JsonPropertyName("statusMessage")]
    public string? StatusMessage { get; set; }

    [JsonPropertyName("currentItem")]
    public int? CurrentItem { get; set; }

    [JsonPropertyName("totalItems")]
    public int? TotalItems { get; set; }
}

             
                                
              
public class PartialOutputPayload
{
    [JsonPropertyName("data")]
    public JsonElement? Data { get; set; }

    [JsonPropertyName("itemIndex")]
    public int ItemIndex { get; set; }
}
