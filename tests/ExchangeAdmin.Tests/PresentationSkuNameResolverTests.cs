using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Tests;

public sealed class PresentationSkuNameResolverTests
{
    private static LocalSkuCatalogDocument BuildDocument(params (string SkuId, string SkuPartNumber, string ProductName)[] entries)
    {
        return new LocalSkuCatalogDocument
        {
            GeneratedOn = "2026-04-14",
            Source = "https://example.com",
            CsvDownload = "https://example.com/sku.csv",
            Entries = entries
                .Select(e => new LocalSkuCatalogEntry
                {
                    SkuId = e.SkuId,
                    SkuPartNumber = e.SkuPartNumber,
                    ProductName = e.ProductName
                })
                .ToList()
        };
    }

    [Fact]
    public void IsLoaded_FalseBeforeReload()
    {
        var resolver = new PresentationSkuNameResolver();
        Assert.False(resolver.IsLoaded);
    }

    [Fact]
    public void IsLoaded_TrueAfterReloadWithEntries()
    {
        var resolver = new PresentationSkuNameResolver();
        resolver.Reload(BuildDocument(("abc-123", "ENTERPRISEPACK", "Microsoft 365 E3")));
        Assert.True(resolver.IsLoaded);
    }

    [Fact]
    public void Resolve_ReturnsFallback_WhenNotLoaded()
    {
        var resolver = new PresentationSkuNameResolver();
        var result = resolver.Resolve("EXCHANGE_S_STANDARD");
        Assert.False(string.IsNullOrWhiteSpace(result));
        Assert.DoesNotContain("EXCHANGE_S_STANDARD", result); // humanised
    }

    [Fact]
    public void Resolve_ReturnsProductName_BySkuPartNumber_AfterReload()
    {
        var resolver = new PresentationSkuNameResolver();
        resolver.Reload(BuildDocument(("abc-123", "ENTERPRISEPACK", "Microsoft 365 E3")));

        Assert.Equal("Microsoft 365 E3", resolver.Resolve("ENTERPRISEPACK"));
    }

    [Fact]
    public void Resolve_IsCaseInsensitive_BySkuPartNumber()
    {
        var resolver = new PresentationSkuNameResolver();
        resolver.Reload(BuildDocument(("abc-123", "ENTERPRISEPACK", "Microsoft 365 E3")));

        Assert.Equal("Microsoft 365 E3", resolver.Resolve("enterprisepack"));
    }

    [Fact]
    public void Resolve_ReturnsProductName_BySkuId_WhenPartNumberMissing()
    {
        var resolver = new PresentationSkuNameResolver();
        resolver.Reload(BuildDocument(("guid-xyz", "SKU_A", "Product A")));

        Assert.Equal("Product A", resolver.Resolve(skuPartNumber: null, skuId: "guid-xyz"));
    }

    [Fact]
    public void Resolve_PrefersPartNumber_OverSkuId()
    {
        var resolver = new PresentationSkuNameResolver();
        resolver.Reload(BuildDocument(
            ("id1", "PART1", "Name By Part"),
            ("id1", "OTHER", "Name By Id")));

        // Both share the same skuId; lookup by part number should win.
        Assert.Equal("Name By Part", resolver.Resolve("PART1", "id1"));
    }

    [Fact]
    public void Reload_ReplacesIndexAtomically()
    {
        var resolver = new PresentationSkuNameResolver();
        resolver.Reload(BuildDocument(("id1", "OLD_SKU", "Old Product")));

        Assert.Equal("Old Product", resolver.Resolve("OLD_SKU"));

        resolver.Reload(BuildDocument(("id2", "NEW_SKU", "New Product")));

        // Old entry is gone; new entry is present.
        var oldResult = resolver.Resolve("OLD_SKU");
        Assert.NotEqual("Old Product", oldResult);
        Assert.Equal("New Product", resolver.Resolve("NEW_SKU"));
    }

    [Fact]
    public void EntryCount_ReflectsLoadedCatalog()
    {
        var resolver = new PresentationSkuNameResolver();
        resolver.Reload(BuildDocument(
            ("id1", "SKU1", "P1"),
            ("id2", "SKU2", "P2"),
            ("id3", "SKU3", "P3")));

        Assert.Equal(3, resolver.EntryCount);
    }

    [Fact]
    public void CatalogVersion_ReflectsGeneratedOn()
    {
        var resolver = new PresentationSkuNameResolver();
        resolver.Reload(BuildDocument(("id1", "SKU1", "P1")));

        Assert.Equal("2026-04-14", resolver.CatalogVersion);
    }

    [Theory]
    [InlineData("EXCHANGE_S_STANDARD", "Exchange S Standard")]
    [InlineData("O365_BUSINESS", "microsoft 365 Business")] // O365 → Microsoft 365
    public void HumanizeFallback_ProducesReadableName(string raw, string expectedContains)
    {
        var result = PresentationSkuNameResolver.HumanizeFallback(raw);
        Assert.Contains(expectedContains, result, StringComparison.OrdinalIgnoreCase);
    }
}
