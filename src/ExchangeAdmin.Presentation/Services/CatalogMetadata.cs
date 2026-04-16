using System.Text.Json.Serialization;

namespace ExchangeAdmin.Presentation.Services;

public sealed class CatalogMetadata
{
    [JsonPropertyName("autoUpdateMode")]
    public string AutoUpdateMode { get; set; } = nameof(CatalogAutoUpdateMode.Daily);

    [JsonPropertyName("lastCheckedUtc")]
    public DateTime? LastCheckedUtc { get; set; }

    [JsonPropertyName("lastUpdatedUtc")]
    public DateTime? LastUpdatedUtc { get; set; }

    [JsonPropertyName("lastError")]
    public string? LastError { get; set; }

    /// <summary>The <c>generatedOn</c> field from the downloaded catalog JSON (e.g. "2026-04-13").</summary>
    [JsonPropertyName("catalogVersion")]
    public string? CatalogVersion { get; set; }

    [JsonPropertyName("entryCount")]
    public int EntryCount { get; set; }
}

[JsonSerializable(typeof(CatalogMetadata))]
internal sealed partial class CatalogMetadataJsonContext : JsonSerializerContext
{
}
