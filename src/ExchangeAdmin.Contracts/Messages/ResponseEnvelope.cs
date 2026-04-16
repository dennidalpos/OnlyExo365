using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Messages;

             
                                         
              
public class ResponseEnvelope : IpcMessage
{
    public ResponseEnvelope()
    {
        Type = MessageType.Response;
    }

    [JsonPropertyName("correlationId")]
    public string CorrelationId { get; set; } = string.Empty;

    [JsonPropertyName("success")]
    public bool Success { get; set; }

    [JsonPropertyName("payload")]
    public JsonElement? Payload { get; set; }

    [JsonPropertyName("error")]
    public NormalizedErrorDto? Error { get; set; }

    [JsonPropertyName("wasCancelled")]
    public bool WasCancelled { get; set; }
}
