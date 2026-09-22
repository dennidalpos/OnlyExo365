using System.Text.Json.Serialization;

namespace OnlyExo365.Shell.Services;

/// <summary>Root document for locally-cached Microsoft 365 SKU catalog JSON.</summary>
public sealed class LocalSkuCatalogDocument
{
    [JsonPropertyName("generatedOn")]
    public string GeneratedOn { get; set; } = string.Empty;

    [JsonPropertyName("source")]
    public string Source { get; set; } = string.Empty;

    [JsonPropertyName("csvDownload")]
    public string CsvDownload { get; set; } = string.Empty;

    [JsonPropertyName("entries")]
    public List<LocalSkuCatalogEntry> Entries { get; set; } = [];
}

public sealed class LocalSkuCatalogEntry
{
    [JsonPropertyName("skuId")]
    public string SkuId { get; set; } = string.Empty;

    [JsonPropertyName("skuPartNumber")]
    public string SkuPartNumber { get; set; } = string.Empty;

    [JsonPropertyName("productName")]
    public string ProductName { get; set; } = string.Empty;

    [JsonPropertyName("servicePlans")]
    public List<LocalSkuCatalogServicePlan> ServicePlans { get; set; } = [];
}

public sealed class LocalSkuCatalogServicePlan
{
    [JsonPropertyName("servicePlanName")]
    public string ServicePlanName { get; set; } = string.Empty;

    [JsonPropertyName("servicePlanId")]
    public string ServicePlanId { get; set; } = string.Empty;

    [JsonPropertyName("friendlyName")]
    public string FriendlyName { get; set; } = string.Empty;
}

[JsonSerializable(typeof(LocalSkuCatalogDocument))]
[JsonSerializable(typeof(LocalSkuCatalogEntry))]
[JsonSerializable(typeof(LocalSkuCatalogServicePlan))]
internal sealed partial class LocalSkuCatalogJsonContext : JsonSerializerContext
{
}

