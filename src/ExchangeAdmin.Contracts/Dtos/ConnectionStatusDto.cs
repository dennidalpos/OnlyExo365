using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Dtos;

/// <summary>
/// Stato connessione Exchange Online.
/// </summary>
public enum ConnectionState
{
    Disconnected,
    Connecting,
    Connected,
    Reconnecting,
    Failed
}

/// <summary>
/// DTO per stato connessione.
/// </summary>
public class ConnectionStatusDto
{
    [JsonPropertyName("state")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public ConnectionState State { get; set; }

    [JsonPropertyName("userPrincipalName")]
    public string? UserPrincipalName { get; set; }

    [JsonPropertyName("organization")]
    public string? Organization { get; set; }

    [JsonPropertyName("connectedAt")]
    public DateTime? ConnectedAt { get; set; }

    [JsonPropertyName("errorMessage")]
    public string? ErrorMessage { get; set; }
}
