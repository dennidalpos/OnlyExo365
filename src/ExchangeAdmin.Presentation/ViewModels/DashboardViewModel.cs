using System.Collections.ObjectModel;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

public class DashboardViewModel : ViewModelBase
{
    private const NavigationPage AlertPage = NavigationPage.Dashboard;
    private readonly IDashboardWorkerService _workerService;
    private readonly NavigationService _navigationService;
    private readonly ShellViewModel _shellViewModel;
    private readonly CacheService _cacheService;

    private CancellationTokenSource? _loadCts;

    private bool _isLoading;
    private DashboardStatsDto? _stats;
    private string? _errorMessage;
    private string? _diagnosticCorrelationId;
    private double _loadingProgress;
    private string? _loadingStatus;
    private int? _loadingCurrentItem;
    private int? _loadingTotalItems;
    private bool _licenseRetryAttempted;

    private static readonly TimeSpan DashboardCacheTtl = TimeSpan.FromMinutes(5);

    public DashboardViewModel(IDashboardWorkerService workerService, NavigationService navigationService, ShellViewModel shellViewModel, CacheService cacheService)
    {
        _workerService = workerService;
        _navigationService = navigationService;
        _shellViewModel = shellViewModel;
        _cacheService = cacheService;

        RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => CanRefresh);
        NavigateToMailboxesCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.Mailboxes));
        NavigateToSharedMailboxesCommand = new RelayCommand(NavigateToSharedMailboxes);
        NavigateToDistributionListsCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.DistributionLists));

        Licenses.CollectionChanged += (_, _) => OnPropertyChanged(nameof(HasLicenses));
        AdminUsers.CollectionChanged += (_, _) => OnPropertyChanged(nameof(HasAdminUsers));
    }

    public bool IsLoading
    {
        get => _isLoading;
        private set
        {
            if (SetProperty(ref _isLoading, value))
            {
                OnPropertyChanged(nameof(CanRefresh));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public DashboardStatsDto? Stats
    {
        get => _stats;
        private set
        {
            if (SetProperty(ref _stats, value))
            {
                OnPropertyChanged(nameof(HasStats));
                OnPropertyChanged(nameof(UserMailboxCount));
                OnPropertyChanged(nameof(SharedMailboxCount));
                OnPropertyChanged(nameof(RoomMailboxCount));
                OnPropertyChanged(nameof(EquipmentMailboxCount));
                OnPropertyChanged(nameof(TotalMailboxCount));
                OnPropertyChanged(nameof(DistributionGroupCount));
                OnPropertyChanged(nameof(DynamicDistributionGroupCount));
                OnPropertyChanged(nameof(UnifiedGroupCount));
                OnPropertyChanged(nameof(TotalGroupCount));
                OnPropertyChanged(nameof(IsLargeTenant));
                OnPropertyChanged(nameof(Warnings));
                OnPropertyChanged(nameof(HasWarnings));
                OnPropertyChanged(nameof(ShowUnifiedGroups));
                OnPropertyChanged(nameof(IsApproximate));
                OnPropertyChanged(nameof(DiagnosticCorrelationId));
                OnPropertyChanged(nameof(HasDiagnosticReference));
                OnPropertyChanged(nameof(DiagnosticReferenceText));
                Licenses.Clear();
                if (value?.Licenses != null)
                {
                    foreach (var lic in value.Licenses)
                    {
                        Licenses.Add(lic);
                    }
                }
                AdminUsers.Clear();
                if (value?.AdminUsers != null)
                {
                    foreach (var admin in value.AdminUsers)
                    {
                        AdminUsers.Add(admin);
                    }
                }

                OnPropertyChanged(nameof(HasLicenses));
                OnPropertyChanged(nameof(HasAdminUsers));
            }
        }
    }

    public bool HasStats => Stats != null;

    public int UserMailboxCount => Stats?.MailboxCounts.UserMailboxes ?? 0;
    public int SharedMailboxCount => Stats?.MailboxCounts.SharedMailboxes ?? 0;
    public int RoomMailboxCount => Stats?.MailboxCounts.RoomMailboxes ?? 0;
    public int EquipmentMailboxCount => Stats?.MailboxCounts.EquipmentMailboxes ?? 0;
    public int TotalMailboxCount => Stats?.MailboxCounts.Total ?? 0;

    public int DistributionGroupCount => Stats?.GroupCounts.DistributionGroups ?? 0;
    public int DynamicDistributionGroupCount => Stats?.GroupCounts.DynamicDistributionGroups ?? 0;
    public int? UnifiedGroupCount => Stats?.GroupCounts.UnifiedGroups;
    public int TotalGroupCount => Stats?.GroupCounts.Total ?? 0;

    public bool IsLargeTenant => Stats?.IsLargeTenant ?? false;
    public List<string> Warnings => Stats?.Warnings ?? new List<string>();
    public bool HasWarnings => Warnings.Count > 0;
    public bool ShowUnifiedGroups => Stats?.GroupCounts.UnifiedGroupsAvailable ?? false;
    public bool IsApproximate => Stats?.MailboxCounts.IsApproximate ?? false;
    public string? DiagnosticCorrelationId => Stats?.CorrelationId ?? _diagnosticCorrelationId;
    public bool HasDiagnosticReference => !string.IsNullOrWhiteSpace(DiagnosticCorrelationId);
    public string? DiagnosticReferenceText => HasDiagnosticReference ? $"Ref: {DiagnosticCorrelationId}" : null;

    public ObservableCollection<TenantLicenseDto> Licenses { get; } = new();
    public bool HasLicenses => Licenses.Count > 0;

    public ObservableCollection<AdminRoleMemberDto> AdminUsers { get; } = new();
    public bool HasAdminUsers => AdminUsers.Count > 0;

    public double LoadingProgress
    {
        get => _loadingProgress;
        private set
        {
            if (SetProperty(ref _loadingProgress, value))
            {
                OnPropertyChanged(nameof(LoadingPercentText));
            }
        }
    }

    public string? LoadingStatus
    {
        get => _loadingStatus;
        private set => SetProperty(ref _loadingStatus, value);
    }

    public int? LoadingCurrentItem
    {
        get => _loadingCurrentItem;
        private set
        {
            if (SetProperty(ref _loadingCurrentItem, value))
            {
                OnPropertyChanged(nameof(HasLoadingItemProgress));
                OnPropertyChanged(nameof(LoadingCountText));
            }
        }
    }

    public int? LoadingTotalItems
    {
        get => _loadingTotalItems;
        private set
        {
            if (SetProperty(ref _loadingTotalItems, value))
            {
                OnPropertyChanged(nameof(HasLoadingItemProgress));
                OnPropertyChanged(nameof(LoadingCountText));
            }
        }
    }

    public bool HasLoadingItemProgress => LoadingCurrentItem.HasValue;
    public string LoadingPercentText => FormatProgressPercent(LoadingProgress);
    public string? LoadingCountText => FormatProgressCount(LoadingCurrentItem, LoadingTotalItems, "items");

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

    public bool CanRefresh => !IsLoading && _shellViewModel.IsExchangeConnected;

    public ICommand RefreshCommand { get; }
    public ICommand NavigateToMailboxesCommand { get; }
    public ICommand NavigateToSharedMailboxesCommand { get; }
    public ICommand NavigateToDistributionListsCommand { get; }

    public async Task LoadAsync()
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            Stats = null;
            ErrorMessage = null;
            _shellViewModel.ClearPageAlert(AlertPage);
            return;
        }

        var cached = _cacheService.Get<DashboardStatsDto>(CacheService.Keys.DashboardStats);
        if (cached != null && cached.Licenses.Count > 0)
        {
            RunOnUiThread(() =>
            {
                Stats = cached;
                ErrorMessage = null;
                _shellViewModel.ClearPageAlert(AlertPage);
            });
            return;
        }

        await RefreshAsync(CancellationToken.None);
    }

    private async Task RefreshAsync(CancellationToken cancellationToken)
    {
        var hasExistingStats = Stats != null;

        _loadCts?.Cancel();
        _loadCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);

        IsLoading = true;
        ErrorMessage = null;
        _shellViewModel.ClearPageAlert(AlertPage);
        _diagnosticCorrelationId = null;
        LoadingProgress = 0;
        LoadingStatus = "Loading dashboard...";
        LoadingCurrentItem = null;
        LoadingTotalItems = null;

        try
        {
            var result = await _workerService.GetDashboardStatsAsync(
                new GetDashboardStatsRequest
                {
                    IncludeUnifiedGroups = true,
                    QuickCount = false
                },
                eventHandler: evt =>
                {
                    if (evt.EventType == EventType.Progress)
                    {
                        var progress = JsonMessageSerializer.ExtractPayload<ProgressEventPayload>(evt.Payload);
                        if (progress != null)
                        {
                            RunOnUiThread(() =>
                            {
                                LoadingProgress = progress.PercentComplete;
                                LoadingStatus = progress.StatusMessage;
                                LoadingCurrentItem = progress.CurrentItem;
                                LoadingTotalItems = progress.TotalItems;
                            });
                        }
                    }
                },
                cancellationToken: _loadCts.Token);

            if (result.IsSuccess && result.Value != null)
            {
                RunOnUiThread(() =>
                {
                    _diagnosticCorrelationId = result.Value.CorrelationId ?? result.CorrelationId;
                    Stats = result.Value;
                    _shellViewModel.ClearPageAlert(AlertPage);
                });
                _cacheService.Set(CacheService.Keys.DashboardStats, result.Value, DashboardCacheTtl);
                await _shellViewModel.RefreshConnectionStatusAsync(_loadCts.Token);

                foreach (var warning in result.Value.Warnings)
                {
                    _shellViewModel.AddLog(LogLevel.Warning, warning, correlationId: DiagnosticCorrelationId);
                }

                if (result.Value.Licenses.Count == 0 && !_licenseRetryAttempted && _shellViewModel.IsExchangeConnected)
                {
                    _licenseRetryAttempted = true;
                    _ = Task.Run(async () =>
                    {
                        await Task.Delay(TimeSpan.FromSeconds(2));
                        if (!_shellViewModel.IsExchangeConnected || IsLoading)
                        {
                            return;
                        }
                        await RefreshAsync(CancellationToken.None);
                    }, CancellationToken.None);
                }
            }
            else if (result.WasCancelled)
            {
            }
            else
            {
                _diagnosticCorrelationId = result.CorrelationId;
                OnPropertyChanged(nameof(DiagnosticCorrelationId));
                OnPropertyChanged(nameof(HasDiagnosticReference));
                OnPropertyChanged(nameof(DiagnosticReferenceText));
                var errorDetails = result.Error != null
                    ? $"{result.Error.Code}: {result.Error.Message}"
                    : "Failed to load dashboard (no error details)";
                if (hasExistingStats)
                {
                    ErrorMessage = errorDetails;
                }
                else
                {
                    _shellViewModel.ShowPageLoadFailedAlert(AlertPage, BuildGlobalAlertDetails(errorDetails));
                }
                _shellViewModel.AddLog(LogLevel.Error, $"Dashboard load failed: {errorDetails}", correlationId: DiagnosticCorrelationId);
            }

            LoadingProgress = 100;
            LoadingStatus = null;
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            var errorDetails = $"Exception: {ex.GetType().Name} - {ex.Message}";
            if (hasExistingStats)
            {
                ErrorMessage = errorDetails;
            }
            else
            {
                _shellViewModel.ShowPageLoadFailedAlert(AlertPage, BuildGlobalAlertDetails(errorDetails));
            }
            _shellViewModel.AddLog(LogLevel.Error, $"Dashboard exception: {ex.GetType().Name} - {ex.Message}", correlationId: DiagnosticCorrelationId);
        }
        finally
        {
            IsLoading = false;
            LoadingProgress = 0;
            LoadingStatus = null;
            LoadingCurrentItem = null;
            LoadingTotalItems = null;
        }
    }

    public void Cancel()
    {
        _loadCts?.Cancel();
    }

    private void NavigateToSharedMailboxes()
    {
        _shellViewModel.Mailboxes?.SetRecipientTypeFilter("SharedMailbox", refresh: false);
        _navigationService.NavigateTo(NavigationPage.Mailboxes);
    }

    private string BuildGlobalAlertDetails(string errorDetails)
    {
        return HasDiagnosticReference
            ? $"{errorDetails}{Environment.NewLine}{Environment.NewLine}{DiagnosticReferenceText}"
            : errorDetails;
    }
}
