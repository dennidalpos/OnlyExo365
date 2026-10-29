using System.Text.Json.Serialization;

namespace OnlyExo365.Contracts.Messages;

             
                                                         
              
public class CancelRequest : IpcMessage
{
    public CancelRequest()
    {
        Type = MessageType.CancelRequest;
    }

    [JsonPropertyName("correlationId")]
    public string CorrelationId { get; set; } = string.Empty;
}

