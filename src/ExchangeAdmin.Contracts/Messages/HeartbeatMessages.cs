using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Messages;

             
                                        
              
public class HeartbeatPing : IpcMessage
{
    public HeartbeatPing()
    {
        Type = MessageType.HeartbeatPing;
    }

    [JsonPropertyName("sequence")]
    public long Sequence { get; set; }
}

             
                                        
              
public class HeartbeatPong : IpcMessage
{
    public HeartbeatPong()
    {
        Type = MessageType.HeartbeatPong;
    }

    [JsonPropertyName("sequence")]
    public long Sequence { get; set; }

    [JsonPropertyName("workerUptime")]
    public TimeSpan WorkerUptime { get; set; }

    [JsonPropertyName("activeOperations")]
    public int ActiveOperations { get; set; }
}
