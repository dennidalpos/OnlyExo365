using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Paging;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

public partial class MigrationViewModel : ViewModelBase
{
    private const NavigationPage AlertPage = NavigationPage.Migration;
    private const int PageSize = PagingDefaults.DefaultPageSize;

    private readonly IMigrationWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;
    private readonly DebounceHelper _searchDebounce = new();

    private bool _isLoading;
    private bool _isLoadingDetails;
    private bool _isApplyingAction;
    private bool _isLoadingEndpoints;
    private bool _isSavingEndpoint;
    private bool _isTestingEndpoint;
    private bool _isRunningPreflight;
    private bool _isCreatingBatch;
    private string? _errorMessage;
    private string? _searchQuery;
    private string _statusFilter = "All";
    private int _totalCount;
    private int _currentSkip;
    private bool _hasMore;
    private MigrationBatchListItemDto? _selectedBatch;
    private MigrationBatchDetailsDto? _selectedDetails;
    private MigrationEndpointDto? _selectedEndpoint;
    private string _endpointName = string.Empty;
    private string _endpointType = "ExchangeRemoteMove";
    private string? _endpointRemoteServer;
    private string? _endpointRpcProxyServer;
    private string? _endpointExchangeServer;
    private string? _endpointEmailAddress;
    private string? _endpointRemoteTenant;
    private int? _endpointPort = 993;
    private string _endpointSecurity = "Ssl";
    private string _endpointAuthentication = "Basic";
    private string? _endpointUsername;
    private string? _pendingEndpointPassword;
    private int _endpointPasswordClearTrigger;
    private int? _endpointMaxConcurrentMigrations = 20;
    private int? _endpointMaxConcurrentIncrementalSyncs = 10;
    private bool _endpointSkipVerification;
    private bool _endpointAcceptUntrustedCertificates;
    private string? _endpointTestSummary;
    private string _newBatchName = string.Empty;
    private string _newBatchType = "Onboarding";
    private string? _newBatchEndpointIdentity;
    private string _newBatchCsvFilePath = string.Empty;
    private string? _newBatchTargetDeliveryDomain;
    private string _newBatchNotificationEmailsText = string.Empty;
    private bool _newBatchAutoStart = true;
    private bool _newBatchAutoComplete;
    private string? _batchPreflightSummary;
    private bool _isBatchPreflightReady;

    public MigrationViewModel(IMigrationWorkerService workerService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;

        RefreshCommand = new AsyncRelayCommand(RefreshAsync, () => CanRefresh);
        LoadMoreCommand = new AsyncRelayCommand(LoadMoreAsync, () => CanLoadMore);
        LoadBatchDetailsCommand = new AsyncRelayCommand(LoadSelectedBatchDetailsAsync, () => CanLoadSelectedBatchDetails);
        StartBatchCommand = new AsyncRelayCommand(StartBatchAsync, () => CanStartSelectedBatch);
        CompleteBatchCommand = new AsyncRelayCommand(CompleteBatchAsync, () => CanCompleteSelectedBatch);
        RemoveBatchCommand = new AsyncRelayCommand(RemoveBatchAsync, () => CanRemoveSelectedBatch);
        RefreshEndpointsCommand = new AsyncRelayCommand(RefreshEndpointsAsync, () => CanRefreshEndpoints);
        SaveEndpointCommand = new AsyncRelayCommand(SaveEndpointAsync, () => CanSaveEndpoint);
        TestEndpointCommand = new AsyncRelayCommand(TestEndpointAsync, () => CanTestEndpoint);
        RunBatchPreflightCommand = new AsyncRelayCommand(RunBatchPreflightAsync, () => CanRunBatchPreflight);
        CreateBatchCommand = new AsyncRelayCommand(CreateBatchAsync, () => CanCreateBatch);
        NewEndpointCommand = new RelayCommand(ResetEndpointEditor, () => _shellViewModel.IsExchangeConnected && !IsBusy);

        _shellViewModel.PropertyChanged += OnShellPropertyChanged;
    }

    public ObservableCollection<MigrationBatchListItemDto> Batches { get; } = new();
    public ObservableCollection<MigrationEndpointDto> Endpoints { get; } = new();
    public CollectionLoadProgressViewModel LoadProgress { get; } = new();

    public IReadOnlyList<string> StatusFilters { get; } =
    [
        "All",
        "Created",
        "Starting",
        "Syncing",
        "Synced",
        "Completing",
        "Completed",
        "CompletedWithErrors",
        "Failed",
        "Stopped",
        "Stopping",
        "Removing"
    ];

    public bool IsLoading
    {
        get => _isLoading;
        private set
        {
            if (SetProperty(ref _isLoading, value))
            {
                OnPropertyChanged(nameof(IsBusy));
                RaiseCanExecuteChanged();
            }
        }
    }

    public bool IsLoadingDetails
    {
        get => _isLoadingDetails;
        private set
        {
            if (SetProperty(ref _isLoadingDetails, value))
            {
                OnPropertyChanged(nameof(IsBusy));
                RaiseCanExecuteChanged();
            }
        }
    }

    public bool IsApplyingAction
    {
        get => _isApplyingAction;
        private set
        {
            if (SetProperty(ref _isApplyingAction, value))
            {
                OnPropertyChanged(nameof(IsBusy));
                RaiseCanExecuteChanged();
            }
        }
    }

    public bool IsLoadingEndpoints
    {
        get => _isLoadingEndpoints;
        private set
        {
            if (SetProperty(ref _isLoadingEndpoints, value))
            {
                OnPropertyChanged(nameof(IsBusy));
                RaiseCanExecuteChanged();
            }
        }
    }

    public bool IsSavingEndpoint
    {
        get => _isSavingEndpoint;
        private set
        {
            if (SetProperty(ref _isSavingEndpoint, value))
            {
                OnPropertyChanged(nameof(IsBusy));
                RaiseCanExecuteChanged();
            }
        }
    }

    public bool IsTestingEndpoint
    {
        get => _isTestingEndpoint;
        private set
        {
            if (SetProperty(ref _isTestingEndpoint, value))
            {
                OnPropertyChanged(nameof(IsBusy));
                RaiseCanExecuteChanged();
            }
        }
    }

    public bool IsRunningPreflight
    {
        get => _isRunningPreflight;
        private set
        {
            if (SetProperty(ref _isRunningPreflight, value))
            {
                OnPropertyChanged(nameof(IsBusy));
                RaiseCanExecuteChanged();
            }
        }
    }

    public bool IsCreatingBatch
    {
        get => _isCreatingBatch;
        private set
        {
            if (SetProperty(ref _isCreatingBatch, value))
            {
                OnPropertyChanged(nameof(IsBusy));
                RaiseCanExecuteChanged();
            }
        }
    }

    public bool IsBusy =>
        IsLoading ||
        IsLoadingDetails ||
        IsApplyingAction ||
        IsLoadingEndpoints ||
        IsSavingEndpoint ||
        IsTestingEndpoint ||
        IsRunningPreflight ||
        IsCreatingBatch;

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

    public bool HasError => !string.IsNullOrWhiteSpace(ErrorMessage);

    public string? SearchQuery
    {
        get => _searchQuery;
        set
        {
            if (SetProperty(ref _searchQuery, value))
            {
                _searchDebounce.Debounce(TriggerRefreshFromUi, 300);
            }
        }
    }

    public string StatusFilter
    {
        get => _statusFilter;
        set
        {
            if (SetProperty(ref _statusFilter, value))
            {
                TriggerRefreshFromUi();
            }
        }
    }

    public int TotalCount
    {
        get => _totalCount;
        private set
        {
            if (SetProperty(ref _totalCount, value))
            {
                OnPropertyChanged(nameof(StatusText));
            }
        }
    }

    public bool HasMore
    {
        get => _hasMore;
        private set
        {
            if (SetProperty(ref _hasMore, value))
            {
                RaiseCanExecuteChanged();
            }
        }
    }

    public string StatusText => $"{Batches.Count} of {TotalCount} batch";

    public MigrationBatchListItemDto? SelectedBatch
    {
        get => _selectedBatch;
        set
        {
            if (SetProperty(ref _selectedBatch, value))
            {
                SelectedDetails = null;
                OnPropertyChanged(nameof(SelectedBatchSummary));
                OnPropertyChanged(nameof(BatchDetailsStatusText));
                OnPropertyChanged(nameof(HasSelectedBatchDetails));
                RaiseCanExecuteChanged();
            }
        }
    }

    public MigrationBatchDetailsDto? SelectedDetails
    {
        get => _selectedDetails;
        private set
        {
            if (SetProperty(ref _selectedDetails, value))
            {
                OnPropertyChanged(nameof(NotificationEmailsText));
                OnPropertyChanged(nameof(CountersText));
                OnPropertyChanged(nameof(HasSelectedBatchDetails));
                OnPropertyChanged(nameof(BatchDetailsStatusText));
            }
        }
    }

    public string SelectedBatchSummary => SelectedBatch == null
        ? "Select a migration batch to review status, counters, and available actions."
        : $"{SelectedBatch.Name} - {SelectedBatch.Status}";

    public bool HasSelectedBatchDetails => SelectedDetails != null;

    public string BatchDetailsStatusText => SelectedBatch switch
    {
        null => "Select a migration batch to enable loading the full details.",
        _ when IsLoadingDetails => "Loading the full migration batch details and report...",
        _ when SelectedDetails != null => "Full details and report loaded on demand.",
        _ => "Full details and report are not preloaded. Use \"Load details\" to fetch them when needed."
    };

    public string NotificationEmailsText => SelectedDetails == null || SelectedDetails.NotificationEmails.Count == 0
        ? "(none)"
        : string.Join("; ", SelectedDetails.NotificationEmails);

    public string CountersText => SelectedDetails == null
        ? "-"
        : $"Total {SelectedDetails.TotalCount ?? 0} | Active {SelectedDetails.ActiveCount ?? 0} | Synced {SelectedDetails.SyncedCount ?? 0} | Finalized {SelectedDetails.FinalizedCount ?? 0} | Failed {SelectedDetails.FailedCount ?? 0} | Stopped {SelectedDetails.StoppedCount ?? 0}";

    public bool CanRefresh => !IsBusy && _shellViewModel.IsExchangeConnected;
    public bool CanLoadMore => !IsBusy && HasMore && _shellViewModel.IsExchangeConnected;
    public bool CanLoadSelectedBatchDetails => CanManageSelectedBatch;
    public bool CanStartSelectedBatch => CanManageSelectedBatch && IsStatusIn(SelectedBatch, "Created", "Stopped", "Failed");
    public bool CanCompleteSelectedBatch => CanManageSelectedBatch && IsStatusIn(SelectedBatch, "Synced", "CompletedWithErrors");
    public bool CanRemoveSelectedBatch => CanManageSelectedBatch;
    public bool CanRefreshEndpoints => !IsBusy && _shellViewModel.IsExchangeConnected;
    public bool CanSaveEndpoint => _shellViewModel.IsExchangeConnected && !IsBusy && HasEndpointEditorMinimumData();
    public bool CanTestEndpoint => _shellViewModel.IsExchangeConnected && !IsBusy && CanTestCurrentEndpoint();
    public bool CanRunBatchPreflight => _shellViewModel.IsExchangeConnected && !IsBusy && HasBatchPreflightMinimumData();
    public bool CanCreateBatch => _shellViewModel.IsExchangeConnected && !IsBusy && IsBatchPreflightReady;

    private bool CanManageSelectedBatch => !IsBusy && _shellViewModel.IsExchangeConnected && SelectedBatch != null;

    public ICommand RefreshCommand { get; }
    public ICommand LoadMoreCommand { get; }
    public ICommand LoadBatchDetailsCommand { get; }
    public ICommand StartBatchCommand { get; }
    public ICommand CompleteBatchCommand { get; }
    public ICommand RemoveBatchCommand { get; }
    public ICommand RefreshEndpointsCommand { get; }
    public ICommand SaveEndpointCommand { get; }
    public ICommand TestEndpointCommand { get; }
    public ICommand RunBatchPreflightCommand { get; }
    public ICommand CreateBatchCommand { get; }
    public ICommand NewEndpointCommand { get; }

    public async Task LoadAsync()
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            ClearStateForDisconnectedSession();
            ErrorMessage = null;
            _shellViewModel.ClearPageAlert(AlertPage);
            return;
        }

        if (Batches.Count == 0)
        {
            await RefreshAsync(CancellationToken.None);
        }

        if (Endpoints.Count == 0)
        {
            await RefreshEndpointsAsync(CancellationToken.None);
        }
    }

    private void TriggerRefreshFromUi()
    {
        _ = SafeRefreshAsync();
    }

    private async Task SafeRefreshAsync()
    {
        try
        {
            await RefreshAsync(CancellationToken.None);
        }
        catch (Exception ex)
        {
            _shellViewModel.AddLog(LogLevel.Error, $"Migration refresh failed: {ex.Message}", "Migration");
        }
    }

    private async Task LoadSelectedBatchDetailsAsync(CancellationToken cancellationToken)
    {
        if (SelectedBatch == null || !_shellViewModel.IsExchangeConnected)
        {
            return;
        }

        try
        {
            await LoadDetailsAsync(SelectedBatch, cancellationToken);
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            _shellViewModel.AddLog(LogLevel.Error, $"Migration details failed: {ex.Message}", "Migration");
        }
    }

    private async Task RefreshAsync(CancellationToken cancellationToken)
    {
        var hasWorkspaceData = HasWorkspaceData;

        if (!_shellViewModel.IsExchangeConnected)
        {
            ClearStateForDisconnectedSession();
            ErrorMessage = null;
            _shellViewModel.ClearPageAlert(AlertPage);
            return;
        }

        IsLoading = true;
        LoadProgress.Start("Loading migration batches...", "batch");
        ErrorMessage = null;
        _shellViewModel.ClearPageAlert(AlertPage);
        var refreshPageSize = GetRefreshPageSize(Batches.Count);
        _currentSkip = 0;
        var previousIdentity = SelectedBatch?.Identity;

        try
        {
            var result = await _workerService.GetMigrationBatchesAsync(
                BuildRequest(0, refreshPageSize),
                eventHandler: HandleLoadProgressEvent,
                cancellationToken: cancellationToken);

            if (!result.IsSuccess || result.Value == null)
            {
                var errorMessage = result.Error?.Message ?? "Unable to load migration batches";
                if (hasWorkspaceData)
                {
                    ErrorMessage = errorMessage;
                }
                else
                {
                    _shellViewModel.ShowPageLoadFailedAlert(AlertPage, errorMessage);
                }
                return;
            }

            Batches.ReplaceAll(result.Value.Batches);
            TotalCount = result.Value.TotalCount;
            HasMore = result.Value.HasMore;
            _currentSkip = Batches.Count;
            _shellViewModel.ClearPageAlert(AlertPage);

            SelectedBatch = RestoreSelection(previousIdentity);
            if (SelectedBatch == null)
            {
                SelectedDetails = null;
            }
        }
        catch (Exception ex)
        {
            if (hasWorkspaceData)
            {
                ErrorMessage = ex.Message;
            }
            else
            {
                _shellViewModel.ShowPageLoadFailedAlert(AlertPage, ex.Message);
            }
        }
        finally
        {
            LoadProgress.Reset();
            IsLoading = false;
        }
    }

    private async Task LoadMoreAsync(CancellationToken cancellationToken)
    {
        if (!HasMore)
        {
            return;
        }

        IsLoading = true;
        LoadProgress.Start("Loading migration batches...", "batch");
        ErrorMessage = null;

        try
        {
            var result = await _workerService.GetMigrationBatchesAsync(
                BuildRequest(_currentSkip),
                eventHandler: HandleLoadProgressEvent,
                cancellationToken: cancellationToken);

            if (!result.IsSuccess || result.Value == null)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to load migration batches";
                return;
            }

            foreach (var item in result.Value.Batches)
            {
                Batches.Add(item);
            }

            TotalCount = result.Value.TotalCount;
            HasMore = result.Value.HasMore;
            _currentSkip = Batches.Count;
            OnPropertyChanged(nameof(StatusText));
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
        finally
        {
            LoadProgress.Reset();
            IsLoading = false;
        }
    }

    private async Task LoadDetailsAsync(MigrationBatchListItemDto batch, CancellationToken cancellationToken)
    {
        IsLoadingDetails = true;
        ErrorMessage = null;

        try
        {
            var result = await _workerService.GetMigrationBatchDetailsAsync(
                new GetMigrationBatchDetailsRequest { Identity = batch.Identity },
                cancellationToken: cancellationToken);

            if (!result.IsSuccess || result.Value == null)
            {
                ErrorMessage = result.Error?.Message ?? "Unable to load migration batch details.";
                return;
            }

            if (SelectedBatch == null || !string.Equals(SelectedBatch.Identity, batch.Identity, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            SelectedDetails = result.Value;
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
        finally
        {
            IsLoadingDetails = false;
        }
    }

    private async Task StartBatchAsync(CancellationToken cancellationToken)
    {
        if (SelectedBatch == null ||
            !ConfirmMutation(
                "Start migration batch",
                SelectedBatch.Identity,
                "Start the selected migration batch.",
                "Confirm migration batch start"))
        {
            return;
        }

        await ExecuteActionAsync(
            "Start",
            new StartMigrationBatchRequest { Identity = SelectedBatch!.Identity },
            (request, token) => _workerService.StartMigrationBatchAsync(request, cancellationToken: token),
            cancellationToken);
    }

    private async Task CompleteBatchAsync(CancellationToken cancellationToken)
    {
        if (SelectedBatch == null ||
            !ConfirmMutation(
                "Complete migration batch",
                SelectedBatch.Identity,
                "Complete the selected migration batch.",
                "Confirm migration batch completion"))
        {
            return;
        }

        await ExecuteActionAsync(
            "Complete",
            new CompleteMigrationBatchRequest { Identity = SelectedBatch!.Identity },
            (request, token) => _workerService.CompleteMigrationBatchAsync(request, cancellationToken: token),
            cancellationToken);
    }

    private async Task RemoveBatchAsync(CancellationToken cancellationToken)
    {
        if (SelectedBatch == null)
        {
            return;
        }

        var confirmed = ErrorDialogService.ShowConfirmation(
            "Confirm migration batch removal",
            $"Operation: remove migration batch\nBatch: {SelectedBatch.Name}\nIdentity: {SelectedBatch.Identity}\n\nConfirm?");

        if (!confirmed)
        {
            return;
        }

        await ExecuteActionAsync(
            "Remove",
            new RemoveMigrationBatchRequest { Identity = SelectedBatch.Identity },
            (request, token) => _workerService.RemoveMigrationBatchAsync(request, cancellationToken: token),
            cancellationToken);
    }

    private async Task ExecuteActionAsync<TRequest>(
        string actionName,
        TRequest request,
        Func<TRequest, CancellationToken, Task<Result>> executor,
        CancellationToken cancellationToken)
    {
        if (SelectedBatch == null)
        {
            return;
        }

        IsApplyingAction = true;
        ErrorMessage = null;

        try
        {
            var result = await executor(request, cancellationToken);
            if (!result.IsSuccess)
            {
                ErrorMessage = result.Error?.Message ?? $"Unable to execute {actionName} on the migration batch.";
                return;
            }

            _shellViewModel.AddLog(LogLevel.Information, $"{actionName} migration batch: {SelectedBatch.Identity}", "Migration");
            await RefreshAsync(cancellationToken);
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
        finally
        {
            IsApplyingAction = false;
        }
    }

    private GetMigrationBatchesRequest BuildRequest(int skip, int? pageSize = null)
    {
        return new GetMigrationBatchesRequest
        {
            SearchQuery = string.IsNullOrWhiteSpace(SearchQuery) ? null : SearchQuery.Trim(),
            Status = NormalizeStatusFilter(StatusFilter),
            PageSize = pageSize ?? PageSize,
            Skip = skip,
            SortBy = "Name"
        };
    }

    private void OnShellPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName != nameof(ShellViewModel.IsExchangeConnected))
        {
            return;
        }

        if (!_shellViewModel.IsExchangeConnected)
        {
            ClearStateForDisconnectedSession();
            ErrorMessage = null;
            _shellViewModel.ClearPageAlert(AlertPage);
        }

        RaiseCanExecuteChanged();
    }

    private void ClearStateForDisconnectedSession()
    {
        Batches.Clear();
        Endpoints.Clear();
        SelectedBatch = null;
        SelectedDetails = null;
        if (SetProperty(ref _selectedEndpoint, null, nameof(SelectedEndpoint)))
        {
            OnPropertyChanged(nameof(IsExistingEndpointSelected));
        }

        TotalCount = 0;
        HasMore = false;
        _currentSkip = 0;
        EndpointTestSummary = null;
        ResetEndpointEditor();
        ResetBatchCreationEditor();
    }

    private MigrationBatchListItemDto? RestoreSelection(string? previousIdentity)
    {
        if (string.IsNullOrWhiteSpace(previousIdentity))
        {
            return Batches.FirstOrDefault();
        }

        return Batches.FirstOrDefault(batch =>
                   string.Equals(batch.Identity, previousIdentity, StringComparison.OrdinalIgnoreCase))
               ?? Batches.FirstOrDefault();
    }

    private void RaiseCanExecuteChanged()
    {
        OnPropertyChanged(nameof(CanRefresh));
        OnPropertyChanged(nameof(CanLoadMore));
        OnPropertyChanged(nameof(CanLoadSelectedBatchDetails));
        OnPropertyChanged(nameof(CanStartSelectedBatch));
        OnPropertyChanged(nameof(CanCompleteSelectedBatch));
        OnPropertyChanged(nameof(CanRemoveSelectedBatch));
        OnPropertyChanged(nameof(CanRefreshEndpoints));
        OnPropertyChanged(nameof(CanSaveEndpoint));
        OnPropertyChanged(nameof(CanTestEndpoint));
        OnPropertyChanged(nameof(CanRunBatchPreflight));
        OnPropertyChanged(nameof(CanCreateBatch));
        CommandManager.InvalidateRequerySuggested();
    }

    private static string NormalizeEndpointType(string? value)
    {
        return value?.Trim() switch
        {
            "ExchangeOutlookAnywhere" => "ExchangeOutlookAnywhere",
            "IMAP" => "IMAP",
            _ => "ExchangeRemoteMove"
        };
    }

    private static string NormalizeBatchType(string? value)
    {
        return value?.Trim() switch
        {
            "Offboarding" => "Offboarding",
            "IMAP" => "IMAP",
            _ => "Onboarding"
        };
    }

    private static bool RequiresCredential(string endpointType)
        => endpointType is "ExchangeRemoteMove" or "ExchangeOutlookAnywhere";

    private static string? NormalizeStatusFilter(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) || string.Equals(value, "All", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return value;
    }

    private static string? TrimToNull(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        return value.Trim();
    }

    private static bool IsStatusIn(MigrationBatchListItemDto? batch, params string[] supportedStatuses)
    {
        if (batch == null || string.IsNullOrWhiteSpace(batch.Status))
        {
            return false;
        }

        return supportedStatuses.Any(status => string.Equals(batch.Status, status, StringComparison.OrdinalIgnoreCase));
    }

    private static int GetRefreshPageSize(int loadedCount)
        => Math.Max(PageSize, loadedCount);

    private void HandleLoadProgressEvent(EventEnvelope evt)
    {
        _ = RunOnUiThreadAsync(() => LoadProgress.Apply(evt));
    }

    private bool HasWorkspaceData =>
        Batches.Count > 0 ||
        Endpoints.Count > 0;
}
