using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Messages;

             
                                                        
              
public class HandshakeRequest : IpcMessage
{
    public HandshakeRequest()
    {
        Type = MessageType.HandshakeRequest;
    }

    [JsonPropertyName("contractsVersion")]
    public string ContractsVersion { get; set; } = ContractVersion.Version;

    [JsonPropertyName("clientId")]
    public string ClientId { get; set; } = Guid.NewGuid().ToString("N");
}
