using System.Text.Json.Serialization;

namespace OnlyExo365.Shell.Services;

/// <summary>
/// Root document for the locally-cached Microsoft 365 SKU catalog JSON.
/// Mirrors the format produced by <c>refresh-microsoft365-sku-catalog.ps1</c>
/// and the embedded Worker resource, but is kept internal to the Presentation
/// layer to avoid coupling to Worker internals.
/// </summary>
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

