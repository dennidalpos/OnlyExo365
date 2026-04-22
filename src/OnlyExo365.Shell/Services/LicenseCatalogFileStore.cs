using System.IO;
using System.Text.Json;

namespace OnlyExo365.Shell.Services;

/// <summary>
/// All file I/O for the local catalog directory.  Writes are atomic:
/// data is written to a <c>.tmp</c> file, validated, then moved over
/// the target with <c>overwrite: true</c>.  A corrupt or partial
/// download can never replace a previously-valid file.
/// </summary>
public sealed class LicenseCatalogFileStore
{
    private readonly string _cacheDirectory;

    private static readonly JsonSerializerOptions JsonReadOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    public LicenseCatalogFileStore(LicenseCatalogConfiguration configuration)
    {
        ArgumentNullException.ThrowIfNull(configuration);
        _cacheDirectory = configuration.ResolveLocalCachePath();
    }

    public string CatalogFilePath =>
        Path.Combine(_cacheDirectory, "Microsoft365SkuCatalog.json");

    public string MetadataFilePath =>
        Path.Combine(_cacheDirectory, "catalog-metadata.json");

    private string CatalogTempFilePath =>
        Path.Combine(_cacheDirectory, "Microsoft365SkuCatalog.json.tmp");

    private string MetadataTempFilePath =>
        Path.Combine(_cacheDirectory, "catalog-metadata.json.tmp");

    public void EnsureDirectoryExists()
    {
        try
        {
            Directory.CreateDirectory(_cacheDirectory);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            throw new CatalogStoreException(
                $"Cannot create catalog directory '{_cacheDirectory}': {ex.Message}", ex);
        }
    }

    public async Task<LocalSkuCatalogDocument?> TryLoadCatalogAsync(CancellationToken cancellationToken = default)
    {
        if (!File.Exists(CatalogFilePath))
        {
            return null;
        }

        try
        {
            await using var stream = File.OpenRead(CatalogFilePath);
            var document = await JsonSerializer.DeserializeAsync(
                stream,
                LocalSkuCatalogJsonContext.Default.LocalSkuCatalogDocument,
                cancellationToken);

            return document?.Entries.Count > 0 ? document : null;
        }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException)
        {
            return null;
        }
    }

    public async Task<CatalogMetadata?> TryLoadMetadataAsync(CancellationToken cancellationToken = default)
    {
        if (!File.Exists(MetadataFilePath))
        {
            return null;
        }

        try
        {
            await using var stream = File.OpenRead(MetadataFilePath);
            return await JsonSerializer.DeserializeAsync(
                stream,
                CatalogMetadataJsonContext.Default.CatalogMetadata,
                cancellationToken);
        }
        catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException)
        {
            return null;
        }
    }

    /// <summary>
    /// Atomically writes a new catalog JSON.  Validates the content can be
    /// deserialised and has at least one entry before replacing the existing file.
    /// </summary>
    public async Task WriteCatalogAtomicAsync(string jsonContent, CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(jsonContent);

        // Validate before touching the live file.
        LocalSkuCatalogDocument? validated;
        try
        {
            validated = JsonSerializer.Deserialize(jsonContent,
                LocalSkuCatalogJsonContext.Default.LocalSkuCatalogDocument);
        }
        catch (JsonException ex)
        {
            throw new CatalogStoreException("Catalog content is not valid JSON: " + ex.Message, ex);
        }

        if (validated == null || validated.Entries.Count == 0)
        {
            throw new CatalogStoreException("Catalog content contains no SKU entries.");
        }

        await WriteAtomicAsync(jsonContent, CatalogTempFilePath, CatalogFilePath, cancellationToken);
    }

    public async Task WriteMetadataAtomicAsync(CatalogMetadata metadata, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(metadata);

        var json = JsonSerializer.Serialize(metadata, CatalogMetadataJsonContext.Default.CatalogMetadata);
        await WriteAtomicAsync(json, MetadataTempFilePath, MetadataFilePath, cancellationToken);
    }

    private static async Task WriteAtomicAsync(
        string content,
        string tempPath,
        string targetPath,
        CancellationToken cancellationToken)
    {
        try
        {
            await File.WriteAllTextAsync(tempPath, content, cancellationToken);
            File.Move(tempPath, targetPath, overwrite: true);
        }
        catch (OperationCanceledException)
        {
            TryDeleteSilently(tempPath);
            throw;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            TryDeleteSilently(tempPath);
            throw new CatalogStoreException(
                $"Failed to write catalog file '{targetPath}': {ex.Message}", ex);
        }
    }

    private static void TryDeleteSilently(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
            // Best-effort cleanup.
        }
    }
}

public sealed class CatalogStoreException : Exception
{
    public CatalogStoreException(string message, Exception? inner = null)
        : base(message, inner)
    {
    }
}

