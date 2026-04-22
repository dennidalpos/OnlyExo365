using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Input;
using OnlyExo365.Shell.Security;
using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts;
using OnlyExo365.Contracts.Diagnostics;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Errors;
using OnlyExo365.Shell.Ipc;
using OnlyExo365.Shell.Localization;
using OnlyExo365.Shell.Text;

namespace OnlyExo365.Shell.ViewModels;

public sealed class ShellViewModel : ViewModelBase, IDisposable
{
    private const int ConnectionAlertPriority = 1000;
    private const int PageAlertPriority = 500;

    private readonly IConnectionWorkerService _workerService;
    private readonly NavigationService _navigationService;
    private readonly ShellConnectionStateViewModel _connectionState;
    private readonly ShellNavigationStateViewModel _navigationState;
    private readonly ShellProgressViewModel _progressState;
    private readonly ShellLogViewModel _logState;
    private readonly ShellPromptViewModel _promptState;
    private readonly Dictionary<string, AlertRegistration> _registeredAlerts = new(StringComparer.Ordinal);

    public LanguageSelectionViewModel? LanguageSelection { get; set; }

    public DashboardViewModel? Dashboard { get; set; }
    public ContactsViewModel? Contacts { get; set; }
    public ResourcesViewModel? Resources { get; set; }
    public PublicFoldersViewModel? PublicFolders { get; set; }
    public MobileDevicesViewModel? MobileDevices { get; set; }
    public MigrationViewModel? Migration { get; set; }
    public PermissionsViewModel? Permissions { get; set; }
    public MailboxListViewModel? Mailboxes { get; set; }
    public DeletedMailboxesViewModel? DeletedMailboxes { get; set; }
    public MailboxDetailsViewModel? MailboxDetails { get; set; }
    public MailboxSpaceViewModel? MailboxSpace { get; set; }
    public MailboxAccessReportViewModel? MailboxAccessReport { get; set; }
    public DistributionListViewModel? DistributionLists { get; set; }
    public MessageTraceViewModel? MessageTrace { get; set; }
    public ComplianceViewModel? Compliance { get; set; }
    public MailSecurityViewModel? MailSecurity { get; set; }
    public MailFlowViewModel? MailFlow { get; set; }
    public ToolsViewModel? Tools { get; set; }
    public LogsViewModel? Logs { get; set; }

    public ShellViewModel(
        IConnectionWorkerService workerService,
        NavigationService navigationService,
        ExchangeOnlineConfiguration? exchangeConfiguration = null,
        IInteractiveExchangeBootstrapService? interactiveExchangeBootstrapService = null)
    {
        _workerService = workerService;
        _navigationService = navigationService;

        _logState = new ShellLogViewModel(new PersistentLogWriter("ui"));
        _progressState = new ShellProgressViewModel();
        _promptState = new ShellPromptViewModel();
        _connectionState = new ShellConnectionStateViewModel(
            workerService,
            _logState,
            exchangeConfiguration,
            interactiveExchangeBootstrapService);
        _navigationState = new ShellNavigationStateViewModel(navigationService, _progressState, _promptState);

        _workerService.StateChanged += OnWorkerStateChanged;
        _workerService.EventReceived += OnEventReceived;
        _workerService.CapabilitiesUpdated += OnCapabilitiesUpdated;

        _connectionState.PropertyChanged += OnSubViewModelPropertyChanged;
        _navigationState.PropertyChanged += OnSubViewModelPropertyChanged;
        _progressState.PropertyChanged += OnSubViewModelPropertyChanged;
        _logState.PropertyChanged += OnSubViewModelPropertyChanged;
        LocalizationService.Instance.CultureChanged += OnCultureChanged;

        UpdateGlobalAlert();
    }

    public NavigationService NavigationService => _navigationService;

    public ShellConnectionStateViewModel Connection => _connectionState;
    public ShellNavigationStateViewModel Navigation => _navigationState;
    public ShellProgressViewModel Progress => _progressState;
    public ShellLogViewModel Logging => _logState;
    public ShellPromptViewModel Prompt => _promptState;
    public AppAlertViewModel GlobalAlert { get; } = new();

    public WorkerConnectionState WorkerState
    {
        get => _connectionState.WorkerState;
        private set => _connectionState.WorkerState = value;
    }

    public string WorkerStateDisplay => _connectionState.WorkerStateDisplay;
    public string WorkerStateColor => _connectionState.WorkerStateColor;
    public bool IsWorkerRunning => _connectionState.IsWorkerRunning;
    public string WorkerRunningDisplay => _connectionState.WorkerRunningDisplay;
    public bool IsWorkerConsoleVisible => _connectionState.IsWorkerConsoleVisible;
    public string WorkerConsoleVisibilityDisplay => _connectionState.WorkerConsoleVisibilityDisplay;
    public bool CanToggleWorkerConsole => _connectionState.CanToggleWorkerConsole;
    public bool IsWorkerConsoleToggleBusy => _connectionState.IsWorkerConsoleToggleBusy;

    public bool IsWorkerBusy
    {
        get => _connectionState.IsWorkerBusy;
        private set => _connectionState.IsWorkerBusy = value;
    }

    public ConnectionState ExchangeState
    {
        get => _connectionState.ExchangeState;
        private set => _connectionState.ExchangeState = value;
    }

    public string ExchangeStateDisplay => _connectionState.ExchangeStateDisplay;
    public string ExchangeStateColor => _connectionState.ExchangeStateColor;
    public bool IsExchangeConnected => _connectionState.IsExchangeConnected;
    public bool IsExchangeConnectionDisabled => _connectionState.IsExchangeConnectionDisabled;

    public bool IsGraphConnected
    {
        get => _connectionState.IsGraphConnected;
        private set => _connectionState.IsGraphConnected = value;
    }

    public string GraphStateColor => _connectionState.GraphStateColor;
    public string GraphStateDisplay => _connectionState.GraphStateDisplay;
    public string GraphStateTooltip => _connectionState.GraphStateTooltip;

    public string? ConnectedUser
    {
        get => _connectionState.ConnectedUser;
        private set => _connectionState.ConnectedUser = value;
    }

    public string? ConnectedOrganization
    {
        get => _connectionState.ConnectedOrganization;
        private set => _connectionState.ConnectedOrganization = value;
    }

    public NavigationPage CurrentPage
    {
        get => _navigationState.CurrentPage;
        private set => _navigationState.CurrentPage = value;
    }

    public bool IsDashboardPage => _navigationState.IsDashboardPage;
    public bool IsContactsPage => _navigationState.IsContactsPage;
    public bool IsResourcesPage => _navigationState.IsResourcesPage;
    public bool IsPublicFoldersPage => _navigationState.IsPublicFoldersPage;
    public bool IsMobileDevicesPage => _navigationState.IsMobileDevicesPage;
    public bool IsMigrationPage => _navigationState.IsMigrationPage;
    public bool IsPermissionsPage => _navigationState.IsPermissionsPage;
    public bool IsMailboxesPage => _navigationState.IsMailboxesPage;
    public bool IsDeletedMailboxesPage => _navigationState.IsDeletedMailboxesPage;
    public bool IsMailboxSpacePage => _navigationState.IsMailboxSpacePage;
    public bool IsMailboxAccessReportPage => _navigationState.IsMailboxAccessReportPage;
    public bool IsDistributionListsPage => _navigationState.IsDistributionListsPage;
    public bool IsMessageTracePage => _navigationState.IsMessageTracePage;
    public bool IsCompliancePage => _navigationState.IsCompliancePage;
    public bool IsMailSecurityPage => _navigationState.IsMailSecurityPage;
    public bool IsMailFlowPage => _navigationState.IsMailFlowPage;
    public bool IsToolsPage => _navigationState.IsToolsPage;
    public bool IsLogsPage => _navigationState.IsLogsPage;
    public string CurrentPageTitle => _navigationState.CurrentPageTitle;

    public CapabilityMapDto? Capabilities
    {
        get => _connectionState.Capabilities;
        private set => _connectionState.Capabilities = value;
    }

    public bool HasCapabilities => _connectionState.HasCapabilities;
    public string CapabilitiesDisplay => _connectionState.CapabilitiesDisplay;
    public bool CanAccessMobileDevicesPage => _connectionState.CanAccessMobileDevicesPage;
    public string MobileDevicesNavigationTooltip => _connectionState.MobileDevicesNavigationTooltip;
    public bool CanAccessMigrationPage => _connectionState.CanAccessMigrationPage;
    public string MigrationNavigationTooltip => _connectionState.MigrationNavigationTooltip;
    public bool CanAccessPermissionsPage => _connectionState.CanAccessPermissionsPage;
    public string PermissionsNavigationTooltip => _connectionState.PermissionsNavigationTooltip;
    public bool CanAccessMessageTracePage => _connectionState.CanAccessMessageTracePage;
    public string MessageTraceNavigationTooltip => _connectionState.MessageTraceNavigationTooltip;
    public bool CanAccessCompliancePage => _connectionState.CanAccessCompliancePage;
    public string ComplianceNavigationTooltip => _connectionState.ComplianceNavigationTooltip;
    public bool CanAccessMailSecurityPage => _connectionState.CanAccessMailSecurityPage;
    public string MailSecurityNavigationTooltip => _connectionState.MailSecurityNavigationTooltip;

    public bool IsGlobalOperationRunning
    {
        get => _progressState.IsGlobalOperationRunning;
        private set => _progressState.IsGlobalOperationRunning = value;
    }

    public int GlobalProgress
    {
        get => _progressState.GlobalProgress;
        set => _progressState.GlobalProgress = value;
    }

    public string? GlobalStatus
    {
        get => _progressState.GlobalStatus;
        set => _progressState.GlobalStatus = value;
    }

    public int? GlobalCurrentItem
    {
        get => _progressState.GlobalCurrentItem;
        private set => _progressState.GlobalCurrentItem = value;
    }

    public int? GlobalTotalItems
    {
        get => _progressState.GlobalTotalItems;
        private set => _progressState.GlobalTotalItems = value;
    }

    public bool HasGlobalItemProgress => _progressState.HasGlobalItemProgress;
    public string GlobalProgressPercentText => _progressState.GlobalProgressPercentText;
    public string? GlobalProgressCountText => _progressState.GlobalProgressCountText;

    public bool IsNavigationLocked => _navigationState.IsNavigationLocked;
    public bool CanNavigate => _navigationState.CanNavigate;

    public bool IsVerboseLoggingEnabled
    {
        get => _logState.IsVerboseLoggingEnabled;
        set => _logState.IsVerboseLoggingEnabled = value;
    }

    public ObservableCollection<LogEntry> LogEntries => _logState.LogEntries;

    public bool CanStartWorker => _connectionState.CanStartWorker;
    public bool CanStopWorker => _connectionState.CanStopWorker;
    public bool CanRestartWorker => _connectionState.CanRestartWorker;
    public bool CanKillWorker => _connectionState.CanKillWorker;
    public bool CanConnectExchange => _connectionState.CanConnectExchange;
    public bool CanDisconnectExchange => _connectionState.CanDisconnectExchange;

    public ICommand StartWorkerCommand => _connectionState.StartWorkerCommand;
    public ICommand StopWorkerCommand => _connectionState.StopWorkerCommand;
    public ICommand RestartWorkerCommand => _connectionState.RestartWorkerCommand;
    public ICommand KillWorkerCommand => _connectionState.KillWorkerCommand;
    public ICommand SetWorkerConsoleVisibilityCommand => _connectionState.SetWorkerConsoleVisibilityCommand;
    public ICommand ConnectExchangeCommand => _connectionState.ConnectExchangeCommand;
    public ICommand DisconnectExchangeCommand => _connectionState.DisconnectExchangeCommand;

    public ICommand NavigateToDashboardCommand => _navigationState.NavigateToDashboardCommand;
    public ICommand NavigateToContactsCommand => _navigationState.NavigateToContactsCommand;
    public ICommand NavigateToResourcesCommand => _navigationState.NavigateToResourcesCommand;
    public ICommand NavigateToPublicFoldersCommand => _navigationState.NavigateToPublicFoldersCommand;
    public ICommand NavigateToMobileDevicesCommand => _navigationState.NavigateToMobileDevicesCommand;
    public ICommand NavigateToMigrationCommand => _navigationState.NavigateToMigrationCommand;
    public ICommand NavigateToPermissionsCommand => _navigationState.NavigateToPermissionsCommand;
    public ICommand NavigateToMailboxesCommand => _navigationState.NavigateToMailboxesCommand;
    public ICommand NavigateToDeletedMailboxesCommand => _navigationState.NavigateToDeletedMailboxesCommand;
    public ICommand NavigateToMailboxSpaceCommand => _navigationState.NavigateToMailboxSpaceCommand;
    public ICommand NavigateToMailboxAccessReportCommand => _navigationState.NavigateToMailboxAccessReportCommand;
    public ICommand NavigateToDistributionListsCommand => _navigationState.NavigateToDistributionListsCommand;
    public ICommand NavigateToMessageTraceCommand => _navigationState.NavigateToMessageTraceCommand;
    public ICommand NavigateToComplianceCommand => _navigationState.NavigateToComplianceCommand;
    public ICommand NavigateToMailSecurityCommand => _navigationState.NavigateToMailSecurityCommand;
    public ICommand NavigateToMailFlowCommand => _navigationState.NavigateToMailFlowCommand;
    public ICommand NavigateToToolsCommand => _navigationState.NavigateToToolsCommand;
    public ICommand NavigateToLogsCommand => _navigationState.NavigateToLogsCommand;

    public ICommand ClearLogsCommand => _logState.ClearLogsCommand;

    public void AddLog(LogLevel level, string message, string? source = null, string? correlationId = null)
    {
        _logState.AddLog(level, message, source, correlationId);
    }

    // -------------------------------------------------------------------------
    // License Catalog relay
    // -------------------------------------------------------------------------

    /// <summary>
    /// Raised when the local SKU catalog is updated or initialised.
    /// ViewModels that display license names can subscribe to this event
    /// and re-normalise their collections without a Worker round-trip.
    /// </summary>
    public event EventHandler<OnlyExo365.Shell.Services.CatalogUpdatedEventArgs>? LicenseCatalogUpdated;

    internal void RaiseLicenseCatalogUpdated(OnlyExo365.Shell.Services.CatalogUpdatedEventArgs args)
        => LicenseCatalogUpdated?.Invoke(this, args);

    public void RegisterBackgroundProgressOperation(string correlationId)
    {
        _progressState.RegisterBackgroundOperation(correlationId);
    }

    public void UnregisterBackgroundProgressOperation(string correlationId)
    {
        _progressState.UnregisterBackgroundOperation(correlationId);
    }

    public async Task RefreshConnectionStatusAsync(CancellationToken cancellationToken = default)
    {
        await _connectionState.RefreshConnectionStatusAsync(cancellationToken);
    }

    public Task StartWorkerOnStartupAsync()
    {
        return _connectionState.StartWorkerOnStartupAsync();
    }

    public void RegisterNavigationStateSource(INotifyPropertyChanged source, Func<bool> isBlocking, params string[] watchedProperties)
    {
        _navigationState.RegisterBlockingStateSource(source, isBlocking, watchedProperties);
    }

    public void RegisterUnsavedChangesCheck(Func<bool> hasUnsavedChanges)
    {
        _navigationState.RegisterUnsavedChangesCheck(hasUnsavedChanges);
    }

    public bool IsFeatureAvailable(Func<FeatureCapabilitiesDto, bool> featureCheck)
    {
        return _connectionState.IsFeatureAvailable(featureCheck);
    }

    public string GetUnavailableTooltip(string featureName)
    {
        return _connectionState.GetUnavailableTooltip(featureName);
    }

    public LeastPrivilegeFeatureEvaluation EvaluateLeastPrivilege(string featureId)
        => _connectionState.EvaluateLeastPrivilege(featureId);

    public void ShowErrorDialog(string title, string message, string? details = null)
    {
        RunOnUiThread(() => ErrorDialogService.ShowError(title, message, details));
    }

    public void ShowErrorDialog(string title, NormalizedErrorDto error)
    {
        RunOnUiThread(() => ErrorDialogService.ShowError(title, error));
    }

    public void ShowErrorDialog(string title, NormalizedError error)
    {
        RunOnUiThread(() => ErrorDialogService.ShowError(title, error));
    }

    public void ShowGlobalAlert(
        string key,
        AppAlertSeverity severity,
        string title,
        string message,
        string? details = null,
        NavigationPage? sourcePage = null,
        int priority = 0)
    {
        RunOnUiThread(() =>
        {
            _registeredAlerts[key] = new AlertRegistration(
                severity,
                title,
                message,
                details,
                sourcePage,
                sourcePage.HasValue ? UiTextCatalog.GetNavigationLabel(sourcePage.Value) : null,
                priority);
            UpdateGlobalAlert();
        });
    }

    public void ClearGlobalAlert(string key)
    {
        RunOnUiThread(() =>
        {
            if (_registeredAlerts.Remove(key))
            {
                UpdateGlobalAlert();
            }
        });
    }

    public void ShowPageLoadFailedAlert(NavigationPage page, string details, AppAlertSeverity severity = AppAlertSeverity.Error)
    {
        ShowGlobalAlert(
            GetPageAlertKey(page),
            severity,
            UserMessageCatalog.LoadFailedAlertTitle,
            UserMessageCatalog.FormatPageUnavailableMessage(UiTextCatalog.GetNavigationLabel(page)),
            details,
            page,
            PageAlertPriority);
    }

    public void ClearPageAlert(NavigationPage page)
    {
        ClearGlobalAlert(GetPageAlertKey(page));
    }

    public void Dispose()
    {
        _workerService.StateChanged -= OnWorkerStateChanged;
        _workerService.EventReceived -= OnEventReceived;
        _workerService.CapabilitiesUpdated -= OnCapabilitiesUpdated;

        _connectionState.PropertyChanged -= OnSubViewModelPropertyChanged;
        _navigationState.PropertyChanged -= OnSubViewModelPropertyChanged;
        _progressState.PropertyChanged -= OnSubViewModelPropertyChanged;
        _logState.PropertyChanged -= OnSubViewModelPropertyChanged;
        LocalizationService.Instance.CultureChanged -= OnCultureChanged;

        _navigationState.Dispose();
    }

    private void OnWorkerStateChanged(object? sender, WorkerConnectionState state)
    {
        RunOnUiThread(() =>
        {
            _connectionState.ApplyWorkerStateChange(state);
            if (state != WorkerConnectionState.Connected)
            {
                _progressState.Reset();
            }
        });
    }

    private void OnEventReceived(object? sender, EventEnvelope evt)
    {
        RunOnUiThread(() =>
        {
            switch (evt.EventType)
            {
                case EventType.Log:
                    var logPayload = JsonMessageSerializer.ExtractPayload<LogEventPayload>(evt.Payload);
                    if (logPayload != null)
                    {
                        _logState.AddLog(logPayload.Level, logPayload.Message, logPayload.Source, evt.CorrelationId);
                    }
                    break;

                case EventType.Progress:
                    var progressPayload = JsonMessageSerializer.ExtractPayload<ProgressEventPayload>(evt.Payload);
                    if (progressPayload != null)
                    {
                        _progressState.Apply(evt.CorrelationId, progressPayload);
                    }
                    break;
            }
        });
    }

    private void OnCapabilitiesUpdated(object? sender, CapabilityMapDto capabilities)
    {
        RunOnUiThread(() => _connectionState.ApplyCapabilities(capabilities));
    }

    private void OnSubViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName is nameof(CurrentPage) or nameof(IsExchangeConnected) or nameof(ExchangeState) or nameof(IsExchangeConnectionDisabled) or nameof(ShellConnectionStateViewModel.HasWorkerStartupAlert) or nameof(ShellConnectionStateViewModel.WorkerStartupAlertDetails))
        {
            UpdateGlobalAlert();
        }

        if (!string.IsNullOrWhiteSpace(e.PropertyName))
        {
            OnPropertyChanged(e.PropertyName);
        }
    }

    private void OnCultureChanged(object? sender, EventArgs e)
    {
        UpdateGlobalAlert();
    }

    private void UpdateGlobalAlert()
    {
        RunOnUiThread(() =>
        {
            if (TryBuildConnectionAlert(out var connectionAlert))
            {
                GlobalAlert.Update(
                    connectionAlert.Severity,
                    connectionAlert.Title,
                    connectionAlert.Message,
                    connectionAlert.Details,
                    connectionAlert.Source,
                    connectionAlert.Priority);
                return;
            }

            var selectedAlert = _registeredAlerts
                .Select(static pair => pair.Value)
                .Where(IsAlertVisibleForCurrentPage)
                .OrderByDescending(alert => alert.Priority)
                .ThenBy(alert => alert.SourcePage.HasValue ? 0 : 1)
                .FirstOrDefault();

            if (selectedAlert == null)
            {
                GlobalAlert.Clear();
                return;
            }

            GlobalAlert.Update(
                selectedAlert.Severity,
                selectedAlert.Title,
                selectedAlert.Message,
                selectedAlert.Details,
                selectedAlert.Source,
                selectedAlert.Priority);
        });
    }

    private bool IsAlertVisibleForCurrentPage(AlertRegistration registration)
    {
        return registration.SourcePage == null || registration.SourcePage == CurrentPage;
    }

    private bool TryBuildConnectionAlert(out AlertRegistration alert)
    {
        if (_connectionState.HasWorkerStartupAlert)
        {
            alert = new AlertRegistration(
                AppAlertSeverity.Error,
                _connectionState.WorkerStartupAlertTitle,
                _connectionState.WorkerStartupAlertMessage,
                _connectionState.WorkerStartupAlertDetails,
                null,
                UiTextCatalog.WorkerLabel,
                ConnectionAlertPriority + 100);
            return true;
        }

        if (IsExchangeConnected)
        {
            alert = default!;
            return false;
        }

        alert = new AlertRegistration(
            AppAlertSeverity.Warning,
            UserMessageCatalog.ConnectionRequiredAlertTitle,
            UserMessageCatalog.ConnectionRequiredAlertMessage,
            null,
            null,
            null,
            ConnectionAlertPriority);
        return true;
    }

    private static string GetPageAlertKey(NavigationPage page)
        => $"page:{page}";

    private sealed record AlertRegistration(
        AppAlertSeverity Severity,
        string Title,
        string Message,
        string? Details,
        NavigationPage? SourcePage,
        string? Source,
        int Priority);
}

