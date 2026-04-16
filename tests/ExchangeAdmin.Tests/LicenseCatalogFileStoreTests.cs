using System.IO;
using System.Text.Json;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Tests;

public sealed class LicenseCatalogFileStoreTests : IDisposable
{
    private readonly string _tempDir = Path.Combine(
        Path.GetTempPath(), "ExchangeAdmin.Tests", Guid.NewGuid().ToString("N"));

    private LicenseCatalogFileStore CreateStore()
    {
        var config = new LicenseCatalogConfiguration { LocalCachePath = _tempDir };
        return new LicenseCatalogFileStore(config);
    }

    private static string MinimalCatalogJson(string generatedOn = "2026-04-14") =>
        $$"""
        {
          "generatedOn": "{{generatedOn}}",
          "source": "https://example.com",
          "csvDownload": "https://example.com/sku.csv",
          "entries": [
            {
              "skuId": "abc-123",
              "skuPartNumber": "ENTERPRISEPACK",
              "productName": "Microsoft 365 E3"
            }
          ]
        }
        """;

    [Fact]
    public void EnsureDirectoryExists_CreatesDirectory()
    {
        var store = CreateStore();
        store.EnsureDirectoryExists();

        Assert.True(Directory.Exists(_tempDir));
    }

    [Fact]
    public async Task TryLoadCatalogAsync_ReturnsNull_WhenFileAbsent()
    {
        var store = CreateStore();
        store.EnsureDirectoryExists();

        var result = await store.TryLoadCatalogAsync();

        Assert.Null(result);
    }

    [Fact]
    public async Task TryLoadCatalogAsync_ReturnsDocument_WhenFileValid()
    {
        var store = CreateStore();
        store.EnsureDirectoryExists();

        await store.WriteCatalogAtomicAsync(MinimalCatalogJson());

        var result = await store.TryLoadCatalogAsync();

        Assert.NotNull(result);
        Assert.Single(result!.Entries);
        Assert.Equal("ENTERPRISEPACK", result.Entries[0].SkuPartNumber);
    }

    [Fact]
    public async Task WriteCatalogAtomicAsync_RemovesTmpFile_AfterCommit()
    {
        var store = CreateStore();
        store.EnsureDirectoryExists();

        await store.WriteCatalogAtomicAsync(MinimalCatalogJson());

        Assert.True(File.Exists(store.CatalogFilePath));
        Assert.False(File.Exists(store.CatalogFilePath + ".tmp"));
    }

    [Fact]
    public async Task WriteCatalogAtomicAsync_ThrowsCatalogStoreException_WhenInvalidJson()
    {
        var store = CreateStore();
        store.EnsureDirectoryExists();

        await Assert.ThrowsAsync<CatalogStoreException>(
            () => store.WriteCatalogAtomicAsync("{ not valid json "));
    }

    [Fact]
    public async Task WriteCatalogAtomicAsync_ThrowsCatalogStoreException_WhenNoEntries()
    {
        var store = CreateStore();
        store.EnsureDirectoryExists();

        const string emptyEntries = """
            {
              "generatedOn": "2026-04-14",
              "source": "",
              "csvDownload": "",
              "entries": []
            }
            """;

        await Assert.ThrowsAsync<CatalogStoreException>(
            () => store.WriteCatalogAtomicAsync(emptyEntries));
    }

    [Fact]
    public async Task WriteCatalogAtomicAsync_DoesNotReplaceExistingFile_OnValidationFailure()
    {
        var store = CreateStore();
        store.EnsureDirectoryExists();

        // Write a valid catalog first.
        await store.WriteCatalogAtomicAsync(MinimalCatalogJson("2026-01-01"));

        // Attempt to overwrite with garbage; should throw and leave old file intact.
        await Assert.ThrowsAsync<CatalogStoreException>(
            () => store.WriteCatalogAtomicAsync("garbage"));

        // Original file must still be readable.
        var loaded = await store.TryLoadCatalogAsync();
        Assert.NotNull(loaded);
        Assert.Equal("2026-01-01", loaded!.GeneratedOn);
    }

    [Fact]
    public async Task WriteMetadataAtomicAsync_RoundTripsAllFields()
    {
        var store = CreateStore();
        store.EnsureDirectoryExists();

        var metadata = new CatalogMetadata
        {
            AutoUpdateMode = "Monthly",
            LastCheckedUtc = new DateTime(2026, 4, 14, 10, 0, 0, DateTimeKind.Utc),
            LastUpdatedUtc = new DateTime(2026, 4, 13, 8, 0, 0, DateTimeKind.Utc),
            LastError = null,
            CatalogVersion = "2026-04-13",
            EntryCount = 1023
        };

        await store.WriteMetadataAtomicAsync(metadata);

        var loaded = await store.TryLoadMetadataAsync();

        Assert.NotNull(loaded);
        Assert.Equal("Monthly", loaded!.AutoUpdateMode);
        Assert.Equal(metadata.LastCheckedUtc, loaded.LastCheckedUtc);
        Assert.Equal(metadata.LastUpdatedUtc, loaded.LastUpdatedUtc);
        Assert.Null(loaded.LastError);
        Assert.Equal("2026-04-13", loaded.CatalogVersion);
        Assert.Equal(1023, loaded.EntryCount);
    }

    [Fact]
    public async Task TryLoadMetadataAsync_ReturnsNull_WhenFileAbsent()
    {
        var store = CreateStore();
        store.EnsureDirectoryExists();

        var result = await store.TryLoadMetadataAsync();

        Assert.Null(result);
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(_tempDir, recursive: true);
        }
        catch
        {
            // Best-effort cleanup.
        }
    }
}
