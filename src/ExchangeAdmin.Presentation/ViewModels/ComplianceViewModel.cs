using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Localization;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

public sealed partial class ComplianceViewModel : ViewModelBase
{
    private const NavigationPage AlertPage = NavigationPage.Compliance;
    private readonly IComplianceWorkerService _workerService;
    private readonly ShellViewModel _shellViewModel;

    private bool _isLoadingWorkspace;
    private bool _isSearchingAudit;
    private bool _isCreatingSearch;
    private bool _isApplyingAction;
    private bool _isAuditQueueProcessing;
    private int _nextAuditTaskSequence = 1;
    private string? _errorMessage;
    private string? _warningsText;
    private string? _diagnosticCorrelationId;
    private bool _isHoldListingUnsupported;
    private string? _holdListingStatusMessage;
    private ComplianceSearchDto? _selectedSearch;
    private ComplianceCaseDto? _selectedCase;
    private ComplianceAuditSearchTaskViewModel? _selectedAuditSearchTask;
    private string _createSearchName = string.Empty;
    private string? _createSearchCaseName;
    private string _createExchangeLocationsText = string.Empty;
    private string? _createContentMatchQuery;
    private string _holdName = string.Empty;
    private string _selectedPurgeType = "SoftDelete";
    private DateTime _auditStartDate = DateTime.Today.AddDays(-1);
    private DateTime _auditEndDate = DateTime.Today;
    private string? _auditUserIdsText;
    private string? _auditOperationsText;
    private string? _auditObjectIdsText;
    private string? _auditFreeText;
    private int _auditMaxResults = 100;

    public ComplianceViewModel(IComplianceWorkerService workerService, ShellViewModel shellViewModel)
    {
        _workerService = workerService;
        _shellViewModel = shellViewModel;

        RefreshWorkspaceCommand = new AsyncRelayCommand(RefreshWorkspaceAsync, () => CanRefreshWorkspace);
        SearchAuditLogCommand = new AsyncRelayCommand(SearchAuditLogAsync, () => CanSearchAuditLog);
        CreateSearchCommand = new AsyncRelayCommand(CreateSearchAsync, () => CanCreateSearch);
        StartSearchCommand = new AsyncRelayCommand(StartSearchAsync, () => CanStartSearch);
        RemoveSearchCommand = new AsyncRelayCommand(RemoveSearchAsync, () => CanRemoveSearch);
        PurgeSearchCommand = new AsyncRelayCommand(PurgeSearchAsync, () => CanPurgeSearch);
        HoldSearchCommand = new AsyncRelayCommand(HoldSearchAsync, () => CanHoldSearch);

        _shellViewModel.PropertyChanged += OnShellPropertyChanged;
    }

    public ObservableCollection<ComplianceSearchDto> Searches { get; } = new();
    public ObservableCollection<ComplianceCaseDto> Cases { get; } = new();
    public ObservableCollection<ComplianceActionSummaryDto> Actions { get; } = new();
    public ObservableCollection<UnifiedAuditLogRecordDto> AuditResults { get; } = new();
    public ObservableCollection<ComplianceAuditSearchTaskViewModel> AuditSearchTasks { get; } = new();

    public IReadOnlyList<string> PurgeTypes { get; } = ["SoftDelete", "HardDelete"];

    public bool IsLoadingWorkspace
    {
        get => _isLoadingWorkspace;
        private set
        {
            if (SetProperty(ref _isLoadingWorkspace, value))
            {
                OnPropertyChanged(nameof(IsBusy));
                OnPropertyChanged(nameof(IsLoadingOverlayVisible));
                OnPropertyChanged(nameof(LoadingOverlayText));
                RaiseCanExecuteChanged();
            }
        }
    }

    public bool IsSearchingAudit
    {
        get => _isSearchingAudit;
        private set
        {
            if (SetProperty(ref _isSearchingAudit, value))
            {
                OnPropertyChanged(nameof(IsLoadingOverlayVisible));
                OnPropertyChanged(nameof(LoadingOverlayText));
                RaiseCanExecuteChanged();
            }
        }
    }

    public bool IsCreatingSearch
    {
        get => _isCreatingSearch;
        private set
        {
            if (SetProperty(ref _isCreatingSearch, value))
            {
                OnPropertyChanged(nameof(IsBusy));
                OnPropertyChanged(nameof(IsLoadingOverlayVisible));
                OnPropertyChanged(nameof(LoadingOverlayText));
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
                OnPropertyChanged(nameof(IsLoadingOverlayVisible));
                OnPropertyChanged(nameof(LoadingOverlayText));
                RaiseCanExecuteChanged();
            }
        }
    }

    public bool IsBusy => IsLoadingWorkspace || IsCreatingSearch || IsApplyingAction;
    public bool IsLoadingOverlayVisible => IsBusy || IsSearchingAudit;
    public string LoadingOverlayText => GetLoadingOverlayText();

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

    public string? WarningsText
    {
        get => _warningsText;
        private set
        {
            if (SetProperty(ref _warningsText, value))
            {
                OnPropertyChanged(nameof(HasWarnings));
            }
        }
    }

    public bool HasWarnings => !string.IsNullOrWhiteSpace(WarningsText);

    public bool IsHoldListingUnsupported
    {
        get => _isHoldListingUnsupported;
        private set
        {
            if (SetProperty(ref _isHoldListingUnsupported, value))
            {
                OnPropertyChanged(nameof(HasHoldListingStatus));
            }
        }
    }

    public string? HoldListingStatusMessage
    {
        get => _holdListingStatusMessage;
        private set
        {
            if (SetProperty(ref _holdListingStatusMessage, value))
            {
                OnPropertyChanged(nameof(HasHoldListingStatus));
            }
        }
    }

    public bool HasHoldListingStatus => IsHoldListingUnsupported && !string.IsNullOrWhiteSpace(HoldListingStatusMessage);

    public string? DiagnosticCorrelationId
    {
        get => _diagnosticCorrelationId;
        private set
        {
            if (SetProperty(ref _diagnosticCorrelationId, value))
            {
                OnPropertyChanged(nameof(HasDiagnosticReference));
                OnPropertyChanged(nameof(DiagnosticReferenceText));
            }
        }
    }

    public bool HasDiagnosticReference => !string.IsNullOrWhiteSpace(DiagnosticCorrelationId);
    public string? DiagnosticReferenceText => HasDiagnosticReference ? $"Ref: {DiagnosticCorrelationId}" : null;

    public ComplianceSearchDto? SelectedSearch
    {
        get => _selectedSearch;
        set
        {
            if (SetProperty(ref _selectedSearch, value))
            {
                if (!string.IsNullOrWhiteSpace(value?.Name))
                {
                    HoldName = $"{value.Name} Hold";
                }

                OnPropertyChanged(nameof(SelectedSearchExchangeLocations));
                OnPropertyChanged(nameof(SelectedSearchQuery));
                RaiseCanExecuteChanged();
            }
        }
    }

    public ComplianceCaseDto? SelectedCase
    {
        get => _selectedCase;
        set
        {
            if (SetProperty(ref _selectedCase, value))
            {
                RaiseCanExecuteChanged();
            }
        }
    }

    public ComplianceAuditSearchTaskViewModel? SelectedAuditSearchTask
    {
        get => _selectedAuditSearchTask;
        set
        {
            if (SetProperty(ref _selectedAuditSearchTask, value))
            {
                SyncSelectedAuditTaskState();
                OnPropertyChanged(nameof(HasSelectedAuditSearchTask));
            }
        }
    }

    public string CreateSearchName
    {
        get => _createSearchName;
        set
        {
            if (SetProperty(ref _createSearchName, value))
            {
                RaiseCanExecuteChanged();
            }
        }
    }

    public string? CreateSearchCaseName
    {
        get => _createSearchCaseName;
        set => SetProperty(ref _createSearchCaseName, value);
    }

    public string CreateExchangeLocationsText
    {
        get => _createExchangeLocationsText;
        set
        {
            if (SetProperty(ref _createExchangeLocationsText, value))
            {
                RaiseCanExecuteChanged();
            }
        }
    }

    public string? CreateContentMatchQuery
    {
        get => _createContentMatchQuery;
        set => SetProperty(ref _createContentMatchQuery, value);
    }

    public string HoldName
    {
        get => _holdName;
        set
        {
            if (SetProperty(ref _holdName, value))
            {
                RaiseCanExecuteChanged();
            }
        }
    }

    public string SelectedPurgeType
    {
        get => _selectedPurgeType;
        set => SetProperty(ref _selectedPurgeType, value);
    }

    public DateTime AuditStartDate
    {
        get => _auditStartDate;
        set
        {
            if (SetProperty(ref _auditStartDate, value))
            {
                RaiseCanExecuteChanged();
            }
        }
    }

    public DateTime AuditEndDate
    {
        get => _auditEndDate;
        set
        {
            if (SetProperty(ref _auditEndDate, value))
            {
                RaiseCanExecuteChanged();
            }
        }
    }

    public string? AuditUserIdsText
    {
        get => _auditUserIdsText;
        set => SetProperty(ref _auditUserIdsText, value);
    }

    public string? AuditOperationsText
    {
        get => _auditOperationsText;
        set => SetProperty(ref _auditOperationsText, value);
    }

    public string? AuditObjectIdsText
    {
        get => _auditObjectIdsText;
        set => SetProperty(ref _auditObjectIdsText, value);
    }

    public string? AuditFreeText
    {
        get => _auditFreeText;
        set => SetProperty(ref _auditFreeText, value);
    }

    public int AuditMaxResults
    {
        get => _auditMaxResults;
        set => SetProperty(ref _auditMaxResults, Math.Clamp(value, 1, 500));
    }

    public string SelectedSearchExchangeLocations => SelectedSearch == null
        ? "-"
        : string.Join(", ", SelectedSearch.ExchangeLocations);

    public string SelectedSearchQuery => string.IsNullOrWhiteSpace(SelectedSearch?.ContentMatchQuery)
        ? "(No query)"
        : SelectedSearch!.ContentMatchQuery!;

    public string SearchesStatusText => $"{Searches.Count} search";
    public string CasesStatusText => $"{Cases.Count} case";
    public string ActionsStatusText => $"{Actions.Count} action";
    public string AuditStatusText => GetAuditStatusText();

    public bool HasAuditSearchTasks => AuditSearchTasks.Count > 0;

    public string AuditTaskCenterStatus => GetAuditTaskCenterStatus();

    public bool CanRefreshWorkspace => _shellViewModel.IsExchangeConnected && !IsBusy;
    public bool CanSearchAuditLog => _shellViewModel.IsExchangeConnected && AuditEndDate.Date >= AuditStartDate.Date;
    public bool CanCreateSearch => _shellViewModel.IsExchangeConnected
        && !IsBusy
        && !string.IsNullOrWhiteSpace(CreateSearchName)
        && ParseMultiValue(CreateExchangeLocationsText).Count > 0;
    public bool CanStartSearch => _shellViewModel.IsExchangeConnected && !IsBusy && SelectedSearch != null;
    public bool CanRemoveSearch => _shellViewModel.IsExchangeConnected && !IsBusy && SelectedSearch != null;
    public bool CanPurgeSearch => _shellViewModel.IsExchangeConnected && !IsBusy && SelectedSearch != null;
    public bool CanHoldSearch => _shellViewModel.IsExchangeConnected
        && !IsBusy
        && SelectedSearch != null
        && SelectedCase != null
        && !string.IsNullOrWhiteSpace(HoldName);

    public ICommand RefreshWorkspaceCommand { get; }
    public ICommand SearchAuditLogCommand { get; }
    public ICommand CreateSearchCommand { get; }
    public ICommand StartSearchCommand { get; }
    public ICommand RemoveSearchCommand { get; }
    public ICommand PurgeSearchCommand { get; }
    public ICommand HoldSearchCommand { get; }

    public async Task LoadAsync()
    {
        if (!_shellViewModel.IsExchangeConnected)
        {
            ClearStateForDisconnectedSession();
            _shellViewModel.ClearPageAlert(AlertPage);
            return;
        }

        if (Searches.Count == 0)
        {
            await RefreshWorkspaceAsync(CancellationToken.None);
        }
    }

    private async Task RefreshWorkspaceAsync(CancellationToken cancellationToken)
    {
        var hasWorkspaceData = HasWorkspaceData;

        if (!_shellViewModel.IsExchangeConnected)
        {
            ClearStateForDisconnectedSession();
            _shellViewModel.ClearPageAlert(AlertPage);
            return;
        }

        IsLoadingWorkspace = true;
        ErrorMessage = null;
        WarningsText = null;
        DiagnosticCorrelationId = null;
        _shellViewModel.ClearPageAlert(AlertPage);
        IsHoldListingUnsupported = false;
        HoldListingStatusMessage = null;

        var selectedSearchName = SelectedSearch?.Name;
        var selectedCaseName = SelectedCase?.Name;

        try
        {
            var result = await _workerService.GetComplianceWorkspaceAsync(
                eventHandler: null,
                cancellationToken: cancellationToken);

            if (!result.IsSuccess || result.Value == null)
            {
                var errorMessage = result.Error?.Message ?? Loc.Get("Compliance.WorkspaceLoadError");
                DiagnosticCorrelationId = result.CorrelationId;
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

            DiagnosticCorrelationId = result.Value.CorrelationId ?? result.CorrelationId;
            IsHoldListingUnsupported = result.Value.IsHoldListingUnsupported;
            HoldListingStatusMessage = result.Value.HoldListingStatusMessage;

            var visibleWarnings = result.Value.Warnings
                .Where(warning => !string.Equals(warning, result.Value.HoldListingStatusMessage, StringComparison.Ordinal))
                .ToList();

            WarningsText = visibleWarnings.Count == 0
                ? null
                : string.Join(Environment.NewLine, visibleWarnings);

            Searches.ReplaceAll(result.Value.Searches);
            Cases.ReplaceAll(result.Value.Cases);
            Actions.ReplaceAll(result.Value.Actions);

            SelectedSearch = Searches.FirstOrDefault(search => string.Equals(search.Name, selectedSearchName, StringComparison.OrdinalIgnoreCase));
            SelectedCase = Cases.FirstOrDefault(item => string.Equals(item.Name, selectedCaseName, StringComparison.OrdinalIgnoreCase));

            OnPropertyChanged(nameof(SearchesStatusText));
            OnPropertyChanged(nameof(CasesStatusText));
            OnPropertyChanged(nameof(ActionsStatusText));

            foreach (var warning in result.Value.Warnings)
            {
                _shellViewModel.AddLog(LogLevel.Warning, warning, "Compliance", DiagnosticCorrelationId);
            }

            _shellViewModel.ClearPageAlert(AlertPage);
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

            _shellViewModel.AddLog(LogLevel.Error, $"Compliance workspace error: {ex.Message}", "Compliance", DiagnosticCorrelationId);
        }
        finally
        {
            IsLoadingWorkspace = false;
        }
    }

    private Task SearchAuditLogAsync(CancellationToken cancellationToken) => EnqueueAuditSearchAsync(cancellationToken);

    private async Task CreateSearchAsync(CancellationToken cancellationToken)
    {
        if (!ConfirmMutation(
                "Creating compliance search",
                CreateSearchName.Trim(),
                "Create a new compliance search with the configured locations and query.",
                "Confirm compliance search creation"))
        {
            return;
        }

        IsCreatingSearch = true;
        ErrorMessage = null;

        try
        {
            var result = await _workerService.CreateComplianceSearchAsync(
                new CreateComplianceSearchRequest
                {
                    Name = CreateSearchName.Trim(),
                    CaseName = string.IsNullOrWhiteSpace(CreateSearchCaseName) ? null : CreateSearchCaseName.Trim(),
                    ExchangeLocations = ParseMultiValue(CreateExchangeLocationsText),
                    ContentMatchQuery = string.IsNullOrWhiteSpace(CreateContentMatchQuery) ? null : CreateContentMatchQuery.Trim()
                },
                cancellationToken: cancellationToken);

            if (!result.IsSuccess)
            {
                ErrorMessage = result.Error?.Message ?? "Creating compliance search failed.";
                return;
            }

            _shellViewModel.AddLog(LogLevel.Information, $"Compliance search created: {CreateSearchName.Trim()}", "Compliance");
            CreateSearchName = string.Empty;
            CreateExchangeLocationsText = string.Empty;
            CreateContentMatchQuery = null;
            await RefreshWorkspaceAsync(cancellationToken);
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
        }
        finally
        {
            IsCreatingSearch = false;
        }
    }

    private async Task StartSearchAsync(CancellationToken cancellationToken)
    {
        if (SelectedSearch == null)
        {
            return;
        }

        if (!ConfirmMutation(
                "Start compliance search",
                SelectedSearch.Name,
                "Start the selected compliance search.",
                "Confirm compliance search start"))
        {
            return;
        }

        IsApplyingAction = true;
        ErrorMessage = null;

        try
        {
            var result = await _workerService.StartComplianceSearchAsync(
                new StartComplianceSearchRequest { Name = SelectedSearch.Name },
                cancellationToken: cancellationToken);

            if (!result.IsSuccess)
            {
                ErrorMessage = result.Error?.Message ?? "Starting the compliance search failed.";
                return;
            }

            _shellViewModel.AddLog(LogLevel.Information, $"Compliance search started: {SelectedSearch.Name}", "Compliance");
            await RefreshWorkspaceAsync(cancellationToken);
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

    private async Task RemoveSearchAsync(CancellationToken cancellationToken)
    {
        if (SelectedSearch == null)
        {
            return;
        }

        var confirmed = ErrorDialogService.ShowConfirmation(
            "Confirm compliance search removal",
            $"Operation: remove compliance search\nSearch: {SelectedSearch.Name}\n\nConfirm?");

        if (!confirmed)
        {
            return;
        }

        IsApplyingAction = true;
        ErrorMessage = null;

        try
        {
            var result = await _workerService.RemoveComplianceSearchAsync(
                new RemoveComplianceSearchRequest { Name = SelectedSearch.Name },
                cancellationToken: cancellationToken);

            if (!result.IsSuccess)
            {
                ErrorMessage = result.Error?.Message ?? "Removing the compliance search failed.";
                return;
            }

            _shellViewModel.AddLog(LogLevel.Warning, $"Compliance search removed: {SelectedSearch.Name}", "Compliance");
            await RefreshWorkspaceAsync(cancellationToken);
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

    private async Task PurgeSearchAsync(CancellationToken cancellationToken)
    {
        if (SelectedSearch == null)
        {
            return;
        }

        var confirmed = ErrorDialogService.ShowConfirmation(
            "Confirm purge",
            $"Operation: purge compliance search\nSearch: {SelectedSearch.Name}\nPurgeType: {SelectedPurgeType}\n\nConfirm?");

        if (!confirmed)
        {
            return;
        }

        await InvokeActionAsync(
            new InvokeComplianceActionRequest
            {
                SearchName = SelectedSearch.Name,
                ActionType = "Purge",
                PurgeType = SelectedPurgeType
            },
            $"Purge requested for {SelectedSearch.Name}",
            cancellationToken);
    }

    private async Task HoldSearchAsync(CancellationToken cancellationToken)
    {
        if (SelectedSearch == null || SelectedCase == null)
        {
            return;
        }

        var confirmed = ErrorDialogService.ShowConfirmation(
            "Confirm hold",
            $"Operation: create hold from compliance search\nSearch: {SelectedSearch.Name}\nCase: {SelectedCase.Name}\nHold: {HoldName}\n\nConfirm?");

        if (!confirmed)
        {
            return;
        }

        await InvokeActionAsync(
            new InvokeComplianceActionRequest
            {
                SearchName = SelectedSearch.Name,
                ActionType = "Hold",
                CaseName = SelectedCase.Name,
                HoldName = HoldName.Trim()
            },
            $"Hold created for {SelectedSearch.Name} in case {SelectedCase.Name}",
            cancellationToken);
    }

    private async Task InvokeActionAsync(InvokeComplianceActionRequest request, string logMessage, CancellationToken cancellationToken)
    {
        IsApplyingAction = true;
        ErrorMessage = null;

        try
        {
            var result = await _workerService.InvokeComplianceActionAsync(request, cancellationToken: cancellationToken);
            if (!result.IsSuccess)
            {
                ErrorMessage = result.Error?.Message ?? "Compliance action failed.";
                return;
            }

            _shellViewModel.AddLog(LogLevel.Information, logMessage, "Compliance");
            await RefreshWorkspaceAsync(cancellationToken);
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

    private void OnShellPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(ShellViewModel.IsExchangeConnected))
        {
            if (!_shellViewModel.IsExchangeConnected)
            {
                ClearStateForDisconnectedSession();
            }

            RaiseCanExecuteChanged();
        }
    }

    private void ClearStateForDisconnectedSession()
    {
        Searches.Clear();
        Cases.Clear();
        Actions.Clear();
        SelectedSearch = null;
        SelectedCase = null;
        WarningsText = null;
        DiagnosticCorrelationId = null;
        IsHoldListingUnsupported = false;
        HoldListingStatusMessage = null;

        ErrorMessage = null;
        OnPropertyChanged(nameof(SearchesStatusText));
        OnPropertyChanged(nameof(CasesStatusText));
        OnPropertyChanged(nameof(ActionsStatusText));
        OnDisconnectedSessionAuditCleanup();
    }

    private void RaiseCanExecuteChanged()
    {
        OnPropertyChanged(nameof(CanRefreshWorkspace));
        OnPropertyChanged(nameof(CanSearchAuditLog));
        OnPropertyChanged(nameof(CanCreateSearch));
        OnPropertyChanged(nameof(CanStartSearch));
        OnPropertyChanged(nameof(CanRemoveSearch));
        OnPropertyChanged(nameof(CanPurgeSearch));
        OnPropertyChanged(nameof(CanHoldSearch));
        CommandManager.InvalidateRequerySuggested();
    }

    partial void OnDisconnectedSessionAuditCleanup();

    private static List<string> ParseMultiValue(string? value)
    {
        return (value ?? string.Empty)
            .Split([',', ';', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(static entry => !string.IsNullOrWhiteSpace(entry))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private string GetLoadingOverlayText()
    {
        if (IsApplyingAction)
        {
            return "Applying Compliance action...";
        }

        if (IsCreatingSearch)
        {
            return "Creating compliance search...";
        }

        if (IsSearchingAudit)
        {
            return "Searching audit logs...";
        }

        return "Loading Compliance workspace...";
    }

    private bool HasWorkspaceData =>
        Searches.Count > 0 ||
        Cases.Count > 0 ||
        Actions.Count > 0;
}

