using System.Collections.Frozen;
using System.Globalization;

namespace OnlyExo365.Shell.Services;

/// <summary>Thread-safe resolver for Microsoft 365 SKU display names with lock-free reads.</summary>
public sealed class PresentationSkuNameResolver
{
    private sealed record CatalogIndex(
        FrozenDictionary<string, string> BySkuPartNumber,
        FrozenDictionary<string, string> BySkuId,
        string? Version,
        int EntryCount)
    {
        public static CatalogIndex Empty { get; } = new(
            FrozenDictionary<string, string>.Empty,
            FrozenDictionary<string, string>.Empty,
            null,
            0);
    }

    private volatile CatalogIndex _index = CatalogIndex.Empty;

    public bool IsLoaded => _index.EntryCount > 0;

    public string? CatalogVersion => _index.Version;

    public int EntryCount => _index.EntryCount;

    /// <summary>Atomically replaces active catalog index from document.</summary>
    public void Reload(LocalSkuCatalogDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);

        var byPartNumber = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var byId = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var entry in document.Entries)
        {
            if (!string.IsNullOrWhiteSpace(entry.SkuPartNumber) &&
                !string.IsNullOrWhiteSpace(entry.ProductName))
            {
                byPartNumber[entry.SkuPartNumber.Trim()] = entry.ProductName.Trim();
            }

            if (!string.IsNullOrWhiteSpace(entry.SkuId) &&
                !string.IsNullOrWhiteSpace(entry.ProductName))
            {
                byId[entry.SkuId.Trim()] = entry.ProductName.Trim();
            }
        }

        _index = new CatalogIndex(
            byPartNumber.ToFrozenDictionary(StringComparer.OrdinalIgnoreCase),
            byId.ToFrozenDictionary(StringComparer.OrdinalIgnoreCase),
            string.IsNullOrWhiteSpace(document.GeneratedOn) ? null : document.GeneratedOn.Trim(),
            document.Entries.Count);
    }

    /// <summary>Resolves product name by SKU part number or ID, with humanized fallback.</summary>
    public string Resolve(string? skuPartNumber, string? skuId = null)
    {
        var idx = _index;

        if (!string.IsNullOrWhiteSpace(skuPartNumber) &&
            idx.BySkuPartNumber.TryGetValue(skuPartNumber.Trim(), out var byPart))
        {
            return byPart;
        }

        if (!string.IsNullOrWhiteSpace(skuId) &&
            idx.BySkuId.TryGetValue(skuId.Trim(), out var byId))
        {
            return byId;
        }

        return HumanizeFallback(!string.IsNullOrWhiteSpace(skuPartNumber) ? skuPartNumber : skuId);
    }

    internal static string HumanizeFallback(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var normalized = value.Trim().Replace('_', ' ');
        normalized = normalized.Replace("O365", "Microsoft 365", StringComparison.OrdinalIgnoreCase);
        normalized = normalized.Replace("M365", "Microsoft 365", StringComparison.OrdinalIgnoreCase);
        return CultureInfo.InvariantCulture.TextInfo.ToTitleCase(normalized.ToLowerInvariant());
    }
}

