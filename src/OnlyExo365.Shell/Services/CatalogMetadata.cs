using System.Text.Json.Serialization;

namespace OnlyExo365.Shell.Services;

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

    /// <summary>Downloaded catalog <c>generatedOn</c> date.</summary>
    [JsonPropertyName("catalogVersion")]
    public string? CatalogVersion { get; set; }

    [JsonPropertyName("entryCount")]
    public int EntryCount { get; set; }
}

[JsonSerializable(typeof(CatalogMetadata))]
internal sealed partial class CatalogMetadataJsonContext : JsonSerializerContext
{
}

