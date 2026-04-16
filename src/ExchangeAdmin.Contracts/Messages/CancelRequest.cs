using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Messages;

             
                                                         
              
public class CancelRequest : IpcMessage
{
    public CancelRequest()
    {
        Type = MessageType.CancelRequest;
    }

    [JsonPropertyName("correlationId")]
    public string CorrelationId { get; set; } = string.Empty;
}
