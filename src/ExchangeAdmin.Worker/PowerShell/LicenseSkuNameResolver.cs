using System.Globalization;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExchangeAdmin.Worker.PowerShell;

internal static class LicenseSkuNameResolver
{
    private const string CatalogResourceName = "ExchangeAdmin.Worker.Data.Microsoft365SkuCatalog.json";
    private static readonly Lazy<SkuCatalogIndex> CatalogIndex = new(LoadCatalogIndex);
    private static readonly string[] ExchangeOnlinePlanNames =
    [
        "EXCHANGE_S_STANDARD",
        "EXCHANGE_S_ENTERPRISE"
    ];

    public static string Resolve(string? skuPartNumber, string? skuId = null)
    {
        var entry = FindEntry(skuPartNumber, skuId);
        if (!string.IsNullOrWhiteSpace(entry?.ProductName))
        {
            return entry.ProductName;
        }

        return HumanizeFallbackName(!string.IsNullOrWhiteSpace(skuPartNumber) ? skuPartNumber : skuId);
    }

    public static IReadOnlyList<LicenseSkuServicePlan> GetServicePlans(string? skuPartNumber, string? skuId = null)
        => FindEntry(skuPartNumber, skuId)?.ServicePlans ?? [];

    public static bool HasExchangeOnlineServicePlan(string? skuPartNumber, string? skuId = null)
        => GetServicePlans(skuPartNumber, skuId).Any(plan =>
            ExchangeOnlinePlanNames.Contains(plan.ServicePlanName, StringComparer.OrdinalIgnoreCase));

    private static SkuCatalogEntry? FindEntry(string? skuPartNumber, string? skuId)
    {
        var catalog = CatalogIndex.Value;

        if (!string.IsNullOrWhiteSpace(skuPartNumber) &&
            catalog.BySkuPartNumber.TryGetValue(skuPartNumber.Trim(), out var bySkuPartNumber))
        {
            return bySkuPartNumber;
        }

        if (!string.IsNullOrWhiteSpace(skuId) &&
            catalog.BySkuId.TryGetValue(skuId.Trim(), out var bySkuId))
        {
            return bySkuId;
        }

        return null;
    }

    private static SkuCatalogIndex LoadCatalogIndex()
    {
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(CatalogResourceName);
        if (stream == null)
        {
            return SkuCatalogIndex.Empty;
        }

        var document = JsonSerializer.Deserialize(stream, LicenseSkuCatalogContext.Default.LicenseSkuCatalogDocument);
        if (document?.Entries == null || document.Entries.Count == 0)
        {
            return SkuCatalogIndex.Empty;
        }

        var bySkuPartNumber = new Dictionary<string, SkuCatalogEntry>(StringComparer.OrdinalIgnoreCase);
        var bySkuId = new Dictionary<string, SkuCatalogEntry>(StringComparer.OrdinalIgnoreCase);

        foreach (var entry in document.Entries)
        {
            if (!string.IsNullOrWhiteSpace(entry.SkuPartNumber))
            {
                bySkuPartNumber[entry.SkuPartNumber] = entry;
            }

            if (!string.IsNullOrWhiteSpace(entry.SkuId))
            {
                bySkuId[entry.SkuId] = entry;
            }
        }

        return new SkuCatalogIndex(bySkuPartNumber, bySkuId);
    }

    private static string HumanizeFallbackName(string? skuPartNumber)
    {
        if (string.IsNullOrWhiteSpace(skuPartNumber))
        {
            return string.Empty;
        }

        var normalized = skuPartNumber.Trim().Replace('_', ' ');
        normalized = normalized.Replace("O365", "Microsoft 365", StringComparison.OrdinalIgnoreCase);
        normalized = normalized.Replace("M365", "Microsoft 365", StringComparison.OrdinalIgnoreCase);

        return CultureInfo.InvariantCulture.TextInfo.ToTitleCase(normalized.ToLowerInvariant());
    }

    private sealed record SkuCatalogIndex(
        IReadOnlyDictionary<string, SkuCatalogEntry> BySkuPartNumber,
        IReadOnlyDictionary<string, SkuCatalogEntry> BySkuId)
    {
        public static SkuCatalogIndex Empty { get; } =
            new(
                new Dictionary<string, SkuCatalogEntry>(StringComparer.OrdinalIgnoreCase),
                new Dictionary<string, SkuCatalogEntry>(StringComparer.OrdinalIgnoreCase));
    }
}

internal sealed class LicenseSkuCatalogDocument
{
    [JsonPropertyName("entries")]
    public List<SkuCatalogEntry> Entries { get; set; } = [];
}

internal sealed class SkuCatalogEntry
{
    [JsonPropertyName("skuId")]
    public string SkuId { get; set; } = string.Empty;

    [JsonPropertyName("skuPartNumber")]
    public string SkuPartNumber { get; set; } = string.Empty;

    [JsonPropertyName("productName")]
    public string ProductName { get; set; } = string.Empty;

    [JsonPropertyName("servicePlans")]
    public List<LicenseSkuServicePlan> ServicePlans { get; set; } = [];
}

internal sealed class LicenseSkuServicePlan
{
    [JsonPropertyName("servicePlanName")]
    public string ServicePlanName { get; set; } = string.Empty;

    [JsonPropertyName("servicePlanId")]
    public string ServicePlanId { get; set; } = string.Empty;

    [JsonPropertyName("friendlyName")]
    public string FriendlyName { get; set; } = string.Empty;
}

[JsonSerializable(typeof(LicenseSkuCatalogDocument))]
internal sealed partial class LicenseSkuCatalogContext : JsonSerializerContext
{
}
