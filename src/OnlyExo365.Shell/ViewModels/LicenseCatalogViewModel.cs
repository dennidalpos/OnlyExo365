using System.Diagnostics;
using System.Globalization;
using System.ComponentModel;
using System.Windows.Input;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Shell.Helpers;
using OnlyExo365.Shell.Localization;
using OnlyExo365.Shell.Services;

namespace OnlyExo365.Shell.ViewModels;

public sealed class LicenseCatalogViewModel : ViewModelBase, IDisposable
{
    private readonly LicenseCatalogUpdateService _catalogService;
    private readonly ShellViewModel _shellViewModel;

    private string _statusSummary = Loc.Get("Tools.Catalog.Loading");
    private string? _catalogVersion;
    private int _entryCount;
    private string? _lastChecked;
    private string? _lastUpdated;
    private string? _errorMessage;
    private bool _isUpdating;
    private CatalogAutoUpdateMode _selectedAutoUpdateMode;
    private CatalogUpdatedEventArgs? _lastCatalogUpdate;

    public LicenseCatalogViewModel(
        LicenseCatalogUpdateService catalogService,
        ShellViewModel shellViewModel)
    {
        ArgumentNullException.ThrowIfNull(catalogService);
        ArgumentNullException.ThrowIfNull(shellViewModel);

        _catalogService = catalogService;
        _shellViewModel = shellViewModel;
        _selectedAutoUpdateMode = _catalogService.Configuration.AutoUpdateMode;

        UpdateNowCommand = new AsyncRelayCommand(
            () => _catalogService.TryCheckAndUpdateAsync(forceDownload: true),
            () => !IsUpdating);

        OpenCatalogFolderCommand = new RelayCommand(OpenCatalogFolder);

        _catalogService.CatalogUpdated += OnCatalogUpdated;
        LocalizationService.Instance.CultureChanged += OnCultureChanged;
    }

    // -------------------------------------------------------------------------
    // Bound properties
    // -------------------------------------------------------------------------

    public string StatusSummary
    {
        get => _statusSummary;
        private set => SetProperty(ref _statusSummary, value);
    }

    public string? CatalogVersion
    {
        get => _catalogVersion;
        private set => SetProperty(ref _catalogVersion, value);
    }

    public int EntryCount
    {
        get => _entryCount;
        private set => SetProperty(ref _entryCount, value);
    }

    public string? LastChecked
    {
        get => _lastChecked;
        private set => SetProperty(ref _lastChecked, value);
    }

    public string? LastUpdated
    {
        get => _lastUpdated;
        private set => SetProperty(ref _lastUpdated, value);
    }

    public string? ErrorMessage
    {
        get => _errorMessage;
        private set
        {
            if (SetProperty(ref _errorMessage, value))
            {
                OnPropertyChanged(nameof(HasError));
            }
        }
    }

    public bool HasError => !string.IsNullOrEmpty(ErrorMessage);

    public bool IsUpdating
    {
        get => _isUpdating;
        private set
        {
            if (SetProperty(ref _isUpdating, value))
            {
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public CatalogAutoUpdateMode SelectedAutoUpdateMode
    {
        get => _selectedAutoUpdateMode;
        set
        {
            if (SetProperty(ref _selectedAutoUpdateMode, value))
            {
                _ = _catalogService.ChangeAutoUpdateModeAsync(value);
                _shellViewModel.AddLog(
                    LogLevel.Information,
                    $"[LicenseCatalog] Auto-update mode changed to: {value}");
            }
        }
    }

    public IReadOnlyList<CatalogAutoUpdateModeOption> AutoUpdateModes { get; } =
        Enum.GetValues<CatalogAutoUpdateMode>()
            .Select(mode => new CatalogAutoUpdateModeOption(mode))
            .ToArray();

    public ICommand UpdateNowCommand { get; }
    public ICommand OpenCatalogFolderCommand { get; }

    // -------------------------------------------------------------------------
    // Event handler
    // -------------------------------------------------------------------------

    private void OnCatalogUpdated(object? sender, CatalogUpdatedEventArgs args)
    {
        RunOnUiThread(() =>
        {
            IsUpdating = false;

            if (args.IsSuccess)
            {
                _lastCatalogUpdate = args;
                ErrorMessage = null;
                CatalogVersion = args.CatalogVersion;
                EntryCount = args.EntryCount;
                LastChecked = FormatUtc(args.LastCheckedUtc);
                LastUpdated = FormatUtc(args.LastUpdatedUtc);
                StatusSummary = BuildStatusSummary(args);

                _shellViewModel.AddLog(
                    LogLevel.Information,
                    $"[LicenseCatalog] Catalog loaded: version={args.CatalogVersion ?? "unknown"}, entries={args.EntryCount}");
            }
            else
            {
                _lastCatalogUpdate = args;
                ErrorMessage = args.Error;
                StatusSummary = Loc.Get("Tools.Catalog.UpdateFailedUsingLastValid");

                _shellViewModel.AddLog(
                    LogLevel.Warning,
                    $"[LicenseCatalog] Update failed: {args.Error}");
            }
        });
    }

    // -------------------------------------------------------------------------
    // Commands
    // -------------------------------------------------------------------------

    private void OpenCatalogFolder()
    {
        try
        {
            var path = _catalogService.Configuration.ResolveLocalCachePath();
            if (!System.IO.Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = path,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Error, $"[LicenseCatalog] Unable to open catalog folder: {ex.Message}");
        }
    }

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    private static string BuildStatusSummary(CatalogUpdatedEventArgs args)
    {
        if (!string.IsNullOrWhiteSpace(args.CatalogVersion))
        {
            return args.EntryCount > 0
                ? Loc.GetFormat("Tools.Catalog.UpToDateWithEntries", args.CatalogVersion, FormatEntryCount(args.EntryCount))
                : Loc.GetFormat("Tools.Catalog.LoadedVersion", args.CatalogVersion);
        }

        return args.EntryCount > 0
            ? Loc.GetFormat("Tools.Catalog.LoadedEntries", FormatEntryCount(args.EntryCount))
            : Loc.Get("Tools.Catalog.NoCatalogAvailable");
    }

    private static string? FormatUtc(DateTime? utc)
    {
        if (utc == null)
        {
            return Loc.Get("Common.Never");
        }

        return utc.Value.ToLocalTime().ToString("yyyy-MM-dd HH:mm");
    }

    private static string FormatEntryCount(int entryCount)
    {
        var culture = CultureInfo.GetCultureInfo(LocalizationService.Instance.CurrentLocale);
        return entryCount.ToString("N0", culture);
    }

    private void OnCultureChanged(object? sender, EventArgs e)
    {
        foreach (var option in AutoUpdateModes)
        {
            option.NotifyCultureChanged();
        }

        if (_lastCatalogUpdate != null)
        {
            LastChecked = FormatUtc(_lastCatalogUpdate.LastCheckedUtc);
            LastUpdated = FormatUtc(_lastCatalogUpdate.LastUpdatedUtc);
            StatusSummary = _lastCatalogUpdate.IsSuccess
                ? BuildStatusSummary(_lastCatalogUpdate)
                : Loc.Get("Tools.Catalog.UpdateFailedUsingLastValid");
        }
    }

    // -------------------------------------------------------------------------
    // IDisposable
    // -------------------------------------------------------------------------

    public void Dispose()
    {
        _catalogService.CatalogUpdated -= OnCatalogUpdated;
        LocalizationService.Instance.CultureChanged -= OnCultureChanged;
    }
}

public sealed class CatalogAutoUpdateModeOption : INotifyPropertyChanged
{
    public CatalogAutoUpdateModeOption(CatalogAutoUpdateMode mode)
    {
        Mode = mode;
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public CatalogAutoUpdateMode Mode { get; }

    public string DisplayName => Mode switch
    {
        CatalogAutoUpdateMode.Disabled => Loc.Get("Tools.AutoUpdate.Disabled"),
        CatalogAutoUpdateMode.Daily => Loc.Get("Tools.AutoUpdate.Daily"),
        CatalogAutoUpdateMode.Monthly => Loc.Get("Tools.AutoUpdate.Monthly"),
        _ => Mode.ToString()
    };

    public void NotifyCultureChanged()
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(DisplayName)));
}

