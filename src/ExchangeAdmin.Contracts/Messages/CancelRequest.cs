using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Messages;

/// <summary>
/// Richiesta di cancellazione di un'operazione in corso.
/// </summary>
public class CancelRequest : IpcMessage
{
    public CancelRequest()
    {
        Type = MessageType.CancelRequest;
    }

    [JsonPropertyName("correlationId")]
    public string CorrelationId { get; set; } = string.Empty;
}
