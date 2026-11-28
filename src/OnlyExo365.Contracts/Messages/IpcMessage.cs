using System.Text.Json.Serialization;

namespace OnlyExo365.Contracts.Messages;

             
                                                                    
              
public class IpcMessage
{
    [JsonPropertyName("type")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public MessageType Type { get; set; }

    [JsonPropertyName("timestamp")]
    public DateTime Timestamp { get; set; } = DateTime.UtcNow;
}

