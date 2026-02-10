using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Messages;

             
                                            
              
public class HandshakeResponse : IpcMessage
{
    public HandshakeResponse()
    {
        Type = MessageType.HandshakeResponse;
    }

    [JsonPropertyName("contractsVersion")]
    public string ContractsVersion { get; set; } = ContractVersion.Version;

    [JsonPropertyName("workerVersion")]
    public string WorkerVersion { get; set; } = "1.0.1";

    [JsonPropertyName("success")]
    public bool Success { get; set; }

    [JsonPropertyName("errorMessage")]
    public string? ErrorMessage { get; set; }

    [JsonPropertyName("isModuleAvailable")]
    public bool IsModuleAvailable { get; set; }

    [JsonPropertyName("powerShellVersion")]
    public string? PowerShellVersion { get; set; }
}
