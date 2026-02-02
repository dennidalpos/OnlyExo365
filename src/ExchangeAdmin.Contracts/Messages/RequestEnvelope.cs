using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Messages;

             
                                         
              
public class RequestEnvelope : IpcMessage
{
    public RequestEnvelope()
    {
        Type = MessageType.Request;
    }

    [JsonPropertyName("correlationId")]
    public string CorrelationId { get; set; } = Guid.NewGuid().ToString("N");

    [JsonPropertyName("operation")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public OperationType Operation { get; set; }

    [JsonPropertyName("payload")]
    public JsonElement? Payload { get; set; }

                 
                                                                   
                  
    [JsonPropertyName("timeoutMs")]
    public int TimeoutMs { get; set; }
}
