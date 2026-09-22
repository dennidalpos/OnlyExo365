using System.IO;
using System.Reflection;

namespace OnlyExo365.Shell.Services;

/// <summary>Manages lifecycle of local Microsoft 365 SKU catalog in Shell process.</summary>
public sealed class LicenseCatalogUpdateService : IDisposable
{
    // Embedded Worker fallback resource
    private const string EmbeddedWorkerCatalogResourceName =
        "OnlyExo365.Worker.Data.Microsoft365SkuCatalog.json";

    private readonly LicenseCatalogConfiguration _configuration;
    private readonly LicenseCatalogFileStore _fileStore;
    private readonly LicenseCatalogDownloader _downloader;
    private readonly PresentationSkuNameResolver _resolver;
    private readonly SemaphoreSlim _updateLock = new(1, 1);

    private Timer? _timer;
    private CatalogMetadata _metadata = new();
    private volatile bool _isUpdating;
    private bool _disposed;

    /// <summary>Raised when catalog is updated; subscribers must marshal to UI thread.</summary>
    public event EventHandler<CatalogUpdatedEventArgs>? CatalogUpdated;

    public LicenseCatalogUpdateService(
        LicenseCatalogConfiguration configuration,
        LicenseCatalogFileStore fileStore,
        LicenseCatalogDownloader downloader,
        PresentationSkuNameResolver resolver)
    {
        ArgumentNullException.ThrowIfNull(configuration);
        ArgumentNullException.ThrowIfNull(fileStore);
        ArgumentNullException.ThrowIfNull(downloader);
        ArgumentNullException.ThrowIfNull(resolver);

        _configuration = configuration;
        _fileStore = fileStore;
        _downloader = downloader;
        _resolver = resolver;
    }

    public bool IsUpdating => _isUpdating;

    public CatalogMetadata Metadata => _metadata;

    public LicenseCatalogConfiguration Configuration => _configuration;

    /// <summary>Initializes service, loads local or embedded catalog, and arms update timer.</summary>
    public async Task InitializeAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            _fileStore.EnsureDirectoryExists();
        }
        catch (CatalogStoreException ex)
        {
            RaiseCatalogUpdated(BuildErrorArgs(ex.Message));
            return;
        }

        var savedMetadata = await _fileStore.TryLoadMetadataAsync(cancellationToken);
        if (savedMetadata != null)
        {
            _metadata = savedMetadata;

            if (Enum.TryParse<CatalogAutoUpdateMode>(_metadata.AutoUpdateMode, out var persistedMode))
            {
                _configuration.AutoUpdateMode = persistedMode;
            }
        }

        var localCatalog = await _fileStore.TryLoadCatalogAsync(cancellationToken);
        if (localCatalog != null)
        {
            _resolver.Reload(localCatalog);
        }
        else
        {
            // Fallback to embedded Worker catalog
            TryLoadEmbeddedWorkerCatalog();
        }

        RaiseCatalogUpdated(BuildSuccessArgs());

        if (_configuration.CheckOnStartup && IsUpdateDue())
        {
            await TryCheckAndUpdateAsync(forceDownload: false, cancellationToken);
        }
        else
        {
            ArmTimer();
        }
    }

    /// <summary>Downloads and applies catalog if update is due or forceDownload is true.</summary>
    public async Task TryCheckAndUpdateAsync(
        bool forceDownload = false,
        CancellationToken cancellationToken = default)
    {
        // Skip if update is already in progress
        if (!await _updateLock.WaitAsync(TimeSpan.Zero, cancellationToken))
        {
            return;
        }

        _isUpdating = true;

        try
        {
            _metadata.LastCheckedUtc = DateTime.UtcNow;
            _metadata.LastError = null;

            var shouldDownload = forceDownload || IsUpdateDue();

            if (shouldDownload)
            {
                await DownloadAndApplyCatalogAsync(cancellationToken);
            }
            else
            {
                await SaveMetadataAsync(cancellationToken);
            }
        }
        catch (OperationCanceledException)
        {
            // Cancelled on shutdown
        }
        catch (Exception ex)
        {
            _metadata.LastError = ex.Message;
            await TrySaveMetadataAsync(cancellationToken);
            RaiseCatalogUpdated(BuildErrorArgs(ex.Message));
        }
        finally
        {
            _isUpdating = false;
            _updateLock.Release();
            ArmTimer();
        }
    }

    /// <summary>Updates auto-update mode preference and re-arms scheduler.</summary>
    public async Task ChangeAutoUpdateModeAsync(
        CatalogAutoUpdateMode mode,
        CancellationToken cancellationToken = default)
    {
        _configuration.AutoUpdateMode = mode;
        _metadata.AutoUpdateMode = mode.ToString();

        await TrySaveMetadataAsync(cancellationToken);
        ArmTimer();
    }

    private async Task DownloadAndApplyCatalogAsync(CancellationToken cancellationToken)
    {
        string json;
        try
        {
            json = await _downloader.DownloadCatalogJsonAsync(cancellationToken);
        }
        catch (CatalogDownloadException ex)
        {
            throw new InvalidOperationException(
                "Catalog download failed: " + ex.Message, ex);
        }

        // Validate before committing
        var incoming = System.Text.Json.JsonSerializer.Deserialize(
            json, LocalSkuCatalogJsonContext.Default.LocalSkuCatalogDocument);

        if (incoming == null || incoming.Entries.Count == 0)
        {
            throw new InvalidOperationException("The downloaded catalog contains no entries.");
        }

        // Skip write if catalog is not newer
        if (!string.IsNullOrWhiteSpace(_metadata.CatalogVersion) &&
            !string.IsNullOrWhiteSpace(incoming.GeneratedOn) &&
            string.Compare(incoming.GeneratedOn, _metadata.CatalogVersion, StringComparison.Ordinal) <= 0)
        {
            await SaveMetadataAsync(cancellationToken);
            RaiseCatalogUpdated(BuildSuccessArgs());
            return;
        }

        await _fileStore.WriteCatalogAtomicAsync(json, cancellationToken);

        _metadata.LastUpdatedUtc = DateTime.UtcNow;
        _metadata.CatalogVersion = incoming.GeneratedOn;
        _metadata.EntryCount = incoming.Entries.Count;

        await SaveMetadataAsync(cancellationToken);

        _resolver.Reload(incoming);
        RaiseCatalogUpdated(BuildSuccessArgs());
    }

    private bool IsUpdateDue()
    {
        if (_configuration.AutoUpdateMode == CatalogAutoUpdateMode.Disabled)
        {
            return false;
        }

        if (_metadata.LastCheckedUtc == null)
        {
            return true;
        }

        var interval = _configuration.AutoUpdateMode == CatalogAutoUpdateMode.Monthly
            ? TimeSpan.FromDays(29)
            : TimeSpan.FromHours(23);

        return DateTime.UtcNow - _metadata.LastCheckedUtc.Value >= interval;
    }

    private void ArmTimer()
    {
        _timer?.Dispose();
        _timer = null;

        if (_configuration.AutoUpdateMode == CatalogAutoUpdateMode.Disabled)
        {
            return;
        }

        var interval = _configuration.AutoUpdateMode == CatalogAutoUpdateMode.Monthly
            ? TimeSpan.FromDays(30)
            : TimeSpan.FromHours(24);

        TimeSpan dueTime;

        if (_metadata.LastCheckedUtc == null)
        {
            dueTime = interval;
        }
        else
        {
            var nextCheck = _metadata.LastCheckedUtc.Value + interval;
            dueTime = nextCheck > DateTime.UtcNow
                ? nextCheck - DateTime.UtcNow
                : TimeSpan.Zero;
        }

        // One-shot timer re-armed inside callback
        _timer = new Timer(
            _ => _ = TimerCallbackAsync(),
            null,
            dueTime,
            Timeout.InfiniteTimeSpan);
    }

    private async Task TimerCallbackAsync()
    {
        if (_disposed)
        {
            return;
        }

        await TryCheckAndUpdateAsync(forceDownload: false);
    }

    private void TryLoadEmbeddedWorkerCatalog()
    {
        try
        {
            // Load Worker embedded catalog resource
            var workerAssembly = TryGetWorkerAssembly();
            if (workerAssembly == null)
            {
                return;
            }

            using var stream = workerAssembly.GetManifestResourceStream(EmbeddedWorkerCatalogResourceName);
            if (stream == null)
            {
                return;
            }

            var document = System.Text.Json.JsonSerializer.Deserialize(
                stream, LocalSkuCatalogJsonContext.Default.LocalSkuCatalogDocument);

            if (document?.Entries.Count > 0)
            {
                _resolver.Reload(document);
                _metadata.CatalogVersion ??= document.GeneratedOn;
                _metadata.EntryCount = document.Entries.Count;
            }
        }
        catch
        {
            // Best-effort fallback
        }
    }

    private static Assembly? TryGetWorkerAssembly()
    {
        var loaded = AppDomain.CurrentDomain.GetAssemblies()
            .FirstOrDefault(a => a.GetName().Name == "OnlyExo365.Worker");
        if (loaded != null)
        {
            return loaded;
        }

        // Load co-located assembly from disk if not already loaded
        var baseDir = AppDomain.CurrentDomain.BaseDirectory;
        var workerPath = Path.Combine(baseDir, "OnlyExo365.Worker.dll");
        if (!File.Exists(workerPath))
        {
            return null;
        }

        try
        {
            return Assembly.LoadFrom(workerPath);
        }
        catch
        {
            return null;
        }
    }

    private async Task SaveMetadataAsync(CancellationToken cancellationToken)
    {
        _metadata.AutoUpdateMode = _configuration.AutoUpdateMode.ToString();
        await _fileStore.WriteMetadataAtomicAsync(_metadata, cancellationToken);
    }

    private async Task TrySaveMetadataAsync(CancellationToken cancellationToken)
    {
        try
        {
            await SaveMetadataAsync(cancellationToken);
        }
        catch
        {
            // Best-effort metadata write
        }
    }

    private CatalogUpdatedEventArgs BuildSuccessArgs() => new()
    {
        IsSuccess = true,
        CatalogVersion = _metadata.CatalogVersion,
        EntryCount = _resolver.EntryCount > 0 ? _resolver.EntryCount : _metadata.EntryCount,
        LastUpdatedUtc = _metadata.LastUpdatedUtc,
        LastCheckedUtc = _metadata.LastCheckedUtc
    };

    private static CatalogUpdatedEventArgs BuildErrorArgs(string message) => new()
    {
        IsSuccess = false,
        Error = message
    };

    private void RaiseCatalogUpdated(CatalogUpdatedEventArgs args)
        => CatalogUpdated?.Invoke(this, args);

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        _timer?.Dispose();
        _timer = null;
        _updateLock.Dispose();
        _downloader.Dispose();
    }
}

