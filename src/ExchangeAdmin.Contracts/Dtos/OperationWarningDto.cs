using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Dtos;

public class OperationWarningDto
{
    [JsonPropertyName("code")]
    public string Code { get; set; } = string.Empty;

    [JsonPropertyName("message")]
    public string Message { get; set; } = string.Empty;

    [JsonPropertyName("scope")]
    public string? Scope { get; set; }

    [JsonPropertyName("itemIdentity")]
    public string? ItemIdentity { get; set; }

    [JsonPropertyName("affectedItemCount")]
    public int? AffectedItemCount { get; set; }

    [JsonPropertyName("sampleItems")]
    public List<string> SampleItems { get; set; } = new();

    [JsonPropertyName("isPartialData")]
    public bool IsPartialData { get; set; }
}
