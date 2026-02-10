using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Input;
using System.Windows.Threading;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Application.UseCases;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Errors;
using ExchangeAdmin.Infrastructure.Ipc;
using ExchangeAdmin.Presentation.Helpers;
using ExchangeAdmin.Presentation.Services;

namespace ExchangeAdmin.Presentation.ViewModels;

public sealed class ShellViewModel : ViewModelBase, IDisposable
{
    private const string DisableExchangeEnvVar = "EXCHANGEADMIN_DISABLE_EXO";
    private readonly IWorkerService _workerService;
    private readonly NavigationService _navigationService;
    private readonly ConnectExchangeUseCase _connectUseCase;

    private WorkerConnectionState _workerState = WorkerConnectionState.NotStarted;
    private bool _isWorkerBusy;

    private ConnectionState _exchangeState = ConnectionState.Disconnected;
    private string? _connectedUser;
    private string? _connectedOrganization;
    private bool _isGraphConnected;

    private NavigationPage _currentPage = NavigationPage.Dashboard;

    private CapabilityMapDto? _capabilities;

    private bool _isGlobalOperationRunning;
    private int _globalProgress;
    private string? _globalStatus;
    private readonly bool _isExchangeConnectionDisabled;
    private bool _isNavigationLocked;
    private readonly List<INotifyPropertyChanged> _navigationStateSources = new();

    private bool _isVerboseLoggingEnabled = false;
    private const int MaxLogEntries = 1000;

    public DashboardViewModel? Dashboard { get; set; }
    public MailboxListViewModel? Mailboxes { get; set; }
    public DeletedMailboxesViewModel? DeletedMailboxes { get; set; }
    public MailboxListViewModel? SharedMailboxes { get; set; }
    public MailboxDetailsViewModel? MailboxDetails { get; set; }
    public MailboxSpaceViewModel? MailboxSpace { get; set; }
    public DistributionListViewModel? DistributionLists { get; set; }
    public MessageTraceViewModel? MessageTrace { get; set; }
    public MailFlowViewModel? MailFlow { get; set; }
    public ToolsViewModel? Tools { get; set; }
    public LogsViewModel? Logs { get; set; }

    public ShellViewModel(IWorkerService workerService, NavigationService navigationService)
    {
        _workerService = workerService;
        _navigationService = navigationService;
        _connectUseCase = new ConnectExchangeUseCase(workerService);
        _isExchangeConnectionDisabled = IsEnvironmentFlagEnabled(DisableExchangeEnvVar);

        _workerService.StateChanged += OnWorkerStateChanged;
        _workerService.EventReceived += OnEventReceived;
        _workerService.CapabilitiesUpdated += OnCapabilitiesUpdated;
        _navigationService.PageChanged += OnPageChanged;
        _navigationService.Navigating += OnNavigating;

        StartWorkerCommand = new AsyncRelayCommand(StartWorkerAsync, () => CanStartWorker);
        StopWorkerCommand = new AsyncRelayCommand(StopWorkerAsync, () => CanStopWorker);
        RestartWorkerCommand = new AsyncRelayCommand(RestartWorkerAsync, () => CanRestartWorker);
        KillWorkerCommand = new RelayCommand(() => _workerService.KillWorker(), () => CanKillWorker);

        ConnectExchangeCommand = new AsyncRelayCommand(ConnectExchangeAsync, () => CanConnectExchange);
        DisconnectExchangeCommand = new AsyncRelayCommand(DisconnectExchangeAsync, () => CanDisconnectExchange);

        NavigateToDashboardCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.Dashboard));
        NavigateToMailboxesCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.Mailboxes));
        NavigateToDeletedMailboxesCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.DeletedMailboxes));
        NavigateToSharedMailboxesCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.SharedMailboxes));
        NavigateToMailboxSpaceCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.MailboxSpace));
        NavigateToDistributionListsCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.DistributionLists));
        NavigateToMessageTraceCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.MessageTrace));
        NavigateToMailFlowCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.MailFlow));
        NavigateToToolsCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.Tools));
        NavigateToLogsCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.Logs));

        ClearLogsCommand = new RelayCommand(() => LogEntries.Clear());
    }

    #region Properties

    public NavigationService NavigationService => _navigationService;

    public WorkerConnectionState WorkerState
    {
        get => _workerState;
        private set
        {
            if (SetProperty(ref _workerState, value))
            {
                NotifyWorkerStatePropertiesChanged();
            }
        }
    }

    public string WorkerStateDisplay => WorkerState switch
    {
        WorkerConnectionState.NotStarted => "Not Started",
        WorkerConnectionState.Starting => "Starting...",
        WorkerConnectionState.WaitingForHandshake => "Initializing...",
        WorkerConnectionState.Connected => "Running",
        WorkerConnectionState.Restarting => "Restarting...",
        WorkerConnectionState.Stopped => "Stopped",
        WorkerConnectionState.Crashed => "Crashed",
        WorkerConnectionState.Unresponsive => "Unresponsive",
        _ => "Unknown"
    };

    public string WorkerStateColor => WorkerState switch
    {
        WorkerConnectionState.Connected => "#4EC9B0",
        WorkerConnectionState.Starting or WorkerConnectionState.WaitingForHandshake or WorkerConnectionState.Restarting => "#DCDCAA",
        WorkerConnectionState.Crashed or WorkerConnectionState.Unresponsive => "#F14C4C",
        _ => "#9D9D9D"
    };

    public bool IsWorkerRunning => WorkerState == WorkerConnectionState.Connected;

    public bool IsWorkerBusy
    {
        get => _isWorkerBusy;
        private set => SetProperty(ref _isWorkerBusy, value);
    }

    public ConnectionState ExchangeState
    {
        get => _exchangeState;
        private set
        {
            if (SetProperty(ref _exchangeState, value))
            {
                NotifyExchangeStatePropertiesChanged();
            }
        }
    }

    public string ExchangeStateDisplay => ExchangeState switch
    {
        _ when IsExchangeConnectionDisabled => "Disabled (policy)",
        ConnectionState.Connected => "Connected",
        ConnectionState.Connecting => "Connecting...",
        ConnectionState.Reconnecting => "Reconnecting...",
        ConnectionState.Failed => "Connection Failed",
        _ => "Disconnected"
    };

    public string ExchangeStateColor => ExchangeState switch
    {
        _ when IsExchangeConnectionDisabled => "#DCDCAA",
        ConnectionState.Connected => "#4EC9B0",
        ConnectionState.Connecting or ConnectionState.Reconnecting => "#DCDCAA",
        ConnectionState.Failed => "#F14C4C",
        _ => "#9D9D9D"
    };

    public bool IsExchangeConnected => ExchangeState == ConnectionState.Connected;
    public bool IsExchangeConnectionDisabled => _isExchangeConnectionDisabled;

    public bool IsGraphConnected
    {
        get => _isGraphConnected;
        private set
        {
            if (SetProperty(ref _isGraphConnected, value))
            {
                OnPropertyChanged(nameof(GraphStateColor));
                OnPropertyChanged(nameof(GraphStateDisplay));
            }
        }
    }

    public string GraphStateColor => IsGraphConnected
        ? "#4EC9B0"
        : "#9D9D9D";

    public string GraphStateDisplay => IsGraphConnected
        ? "Connected"
        : "Disconnected";

    public string? ConnectedUser
    {
        get => _connectedUser;
        private set
        {
            if (SetProperty(ref _connectedUser, value))
            {
                OnPropertyChanged(nameof(ExchangeStateDisplay));
            }
        }
    }

    public string? ConnectedOrganization
    {
        get => _connectedOrganization;
        private set => SetProperty(ref _connectedOrganization, value);
    }

    public NavigationPage CurrentPage
    {
        get => _currentPage;
        private set
        {
            if (SetProperty(ref _currentPage, value))
            {
                NotifyNavigationPropertiesChanged();
            }
        }
    }

    public bool IsDashboardPage => CurrentPage == NavigationPage.Dashboard;
    public bool IsMailboxesPage => CurrentPage == NavigationPage.Mailboxes;
    public bool IsDeletedMailboxesPage => CurrentPage == NavigationPage.DeletedMailboxes;
    public bool IsSharedMailboxesPage => CurrentPage == NavigationPage.SharedMailboxes;
    public bool IsMailboxSpacePage => CurrentPage == NavigationPage.MailboxSpace;
    public bool IsDistributionListsPage => CurrentPage == NavigationPage.DistributionLists;
    public bool IsMessageTracePage => CurrentPage == NavigationPage.MessageTrace;
    public bool IsMailFlowPage => CurrentPage == NavigationPage.MailFlow;
    public bool IsToolsPage => CurrentPage == NavigationPage.Tools;
    public bool IsLogsPage => CurrentPage == NavigationPage.Logs;

    public string CurrentPageTitle => CurrentPage switch
    {
        NavigationPage.Dashboard => "Dashboard",
        NavigationPage.Mailboxes => "Mailboxes",
        NavigationPage.DeletedMailboxes => "Deleted Mailbox",
        NavigationPage.SharedMailboxes => "Shared Mailboxes",
        NavigationPage.MailboxSpace => "Spazio mailbox",
        NavigationPage.DistributionLists => "Distribution Lists",
        NavigationPage.MessageTrace => "Traccia Messaggi",
        NavigationPage.MailFlow => "Mail Flow",
        NavigationPage.Tools => "Tools",
        NavigationPage.Logs => "Logs",
        _ => "Exchange Admin"
    };

    public CapabilityMapDto? Capabilities
    {
        get => _capabilities;
        private set
        {
            if (SetProperty(ref _capabilities, value))
            {
                OnPropertyChanged(nameof(HasCapabilities));
            }
        }
    }

    public bool HasCapabilities => Capabilities != null;

    public bool IsGlobalOperationRunning
    {
        get => _isGlobalOperationRunning;
        private set
        {
            if (SetProperty(ref _isGlobalOperationRunning, value))
            {
                UpdateNavigationLock();
            }
        }
    }

    public int GlobalProgress
    {
        get => _globalProgress;
        set => SetProperty(ref _globalProgress, value);
    }

    public string? GlobalStatus
    {
        get => _globalStatus;
        set => SetProperty(ref _globalStatus, value);
    }

    public bool IsNavigationLocked
    {
        get => _isNavigationLocked;
        private set
        {
            if (SetProperty(ref _isNavigationLocked, value))
            {
                OnPropertyChanged(nameof(CanNavigate));
            }
        }
    }

    public bool CanNavigate => !IsNavigationLocked;

    public bool IsVerboseLoggingEnabled
    {
        get => _isVerboseLoggingEnabled;
        set
        {
            if (SetProperty(ref _isVerboseLoggingEnabled, value))
            {
                AddLog(value ? LogLevel.Information : LogLevel.Information,
                       value ? "Verbose logging enabled - all PowerShell output will be shown"
                             : "Verbose logging disabled - only important messages will be shown");
            }
        }
    }

    public ObservableCollection<LogEntry> LogEntries { get; } = new();

    #endregion

    #region Can Execute

    public bool CanStartWorker => WorkerState == WorkerConnectionState.NotStarted ||
                                   WorkerState == WorkerConnectionState.Stopped ||
                                   WorkerState == WorkerConnectionState.Crashed;

    public bool CanStopWorker => WorkerState == WorkerConnectionState.Connected;

    public bool CanRestartWorker => WorkerState == WorkerConnectionState.Connected ||
                                     WorkerState == WorkerConnectionState.Crashed ||
                                     WorkerState == WorkerConnectionState.Unresponsive;

    public bool CanKillWorker => WorkerState != WorkerConnectionState.NotStarted &&
                                  WorkerState != WorkerConnectionState.Stopped;

    public bool CanConnectExchange => WorkerState == WorkerConnectionState.Connected &&
                                       ExchangeState == ConnectionState.Disconnected &&
                                       !IsExchangeConnectionDisabled;

    public bool CanDisconnectExchange => WorkerState == WorkerConnectionState.Connected &&
                                          ExchangeState == ConnectionState.Connected;

    #endregion

    #region Commands

    public ICommand StartWorkerCommand { get; }
    public ICommand StopWorkerCommand { get; }
    public ICommand RestartWorkerCommand { get; }
    public ICommand KillWorkerCommand { get; }
    public ICommand ConnectExchangeCommand { get; }
    public ICommand DisconnectExchangeCommand { get; }

    public ICommand NavigateToDashboardCommand { get; }
    public ICommand NavigateToMailboxesCommand { get; }
    public ICommand NavigateToDeletedMailboxesCommand { get; }
    public ICommand NavigateToSharedMailboxesCommand { get; }
    public ICommand NavigateToMailboxSpaceCommand { get; }
    public ICommand NavigateToDistributionListsCommand { get; }
    public ICommand NavigateToMessageTraceCommand { get; }
    public ICommand NavigateToMailFlowCommand { get; }
    public ICommand NavigateToToolsCommand { get; }
    public ICommand NavigateToLogsCommand { get; }

    public ICommand ClearLogsCommand { get; }

    #endregion

    #region Command Implementations

    private async Task StartWorkerAsync(CancellationToken cancellationToken)
    {
        AddLog(LogLevel.Information, "Starting worker...");

        var success = await _workerService.StartWorkerAsync(cancellationToken);

        if (success)
        {
            AddLog(LogLevel.Information, "Worker started successfully");

            var status = _workerService.Status;
            if (!status.IsModuleAvailable)
            {
                var warning = "ExchangeOnlineManagement module is not available.\n\n" +
                            "To use this application, you need to install the module:\n" +
                            "1. Open PowerShell 7 as Administrator\n" +
                            "2. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser";
                AddLog(LogLevel.Warning, "ExchangeOnlineManagement module is not available. Install it with: Install-Module ExchangeOnlineManagement");
                ErrorDialogService.ShowWarning("Module Not Found", warning);
            }
        }
        else
        {
            var errorMsg = _workerService.Status.LastError ?? "Unknown error";
            AddLog(LogLevel.Error, $"Failed to start worker: {errorMsg}");

            ShowErrorDialog("Worker Failed to Start",
                "The background worker process could not be started.",
                $"Error: {errorMsg}\n\n" +
                "Please ensure:\n" +
                "1. PowerShell 7+ is installed\n" +
                "2. The worker executable exists in the application folder\n" +
                "3. You have permission to run executables");
        }
    }

    private async Task StopWorkerAsync(CancellationToken cancellationToken)
    {
        AddLog(LogLevel.Information, "Stopping worker...");
        await _workerService.StopWorkerAsync();
        AddLog(LogLevel.Information, "Worker stopped");
    }

    private async Task RestartWorkerAsync(CancellationToken cancellationToken)
    {
        AddLog(LogLevel.Information, "Restarting worker...");
        var success = await _workerService.RestartWorkerAsync(cancellationToken);

        if (success)
        {
            AddLog(LogLevel.Information, "Worker restarted successfully");
        }
        else
        {
            AddLog(LogLevel.Error, $"Failed to restart worker: {_workerService.Status.LastError}");
        }
    }

    private async Task ConnectExchangeAsync(CancellationToken cancellationToken)
    {
        if (IsExchangeConnectionDisabled)
        {
            AddLog(LogLevel.Warning, "Exchange Online connections are disabled by policy (EXCHANGEADMIN_DISABLE_EXO=1).");
            ErrorDialogService.ShowWarning(
                "Connection Disabled",
                "Exchange Online connections are disabled by policy.\n\n" +
                "To enable connections, unset EXCHANGEADMIN_DISABLE_EXO and restart the application.");
            return;
        }

        ExchangeState = ConnectionState.Connecting;
        AddLog(LogLevel.Information, "Connecting to Exchange Online...");

        var result = await _connectUseCase.ExecuteAsync(
            onLog: (level, msg) => RunOnUiThread(() => AddLog(level, msg)),
            cancellationToken: cancellationToken);

        if (result.IsSuccess && result.Value != null)
        {
            ExchangeState = result.Value.State;
            ConnectedUser = result.Value.UserPrincipalName;
            ConnectedOrganization = result.Value.Organization;
            IsGraphConnected = result.Value.GraphConnected;
            AddLog(LogLevel.Information, $"Connected as {ConnectedUser} to {ConnectedOrganization}");
        }
        else if (result.WasCancelled)
        {
            ExchangeState = ConnectionState.Disconnected;
            IsGraphConnected = false;
            AddLog(LogLevel.Warning, "Connection cancelled");
        }
        else
        {
            ExchangeState = ConnectionState.Failed;
            IsGraphConnected = false;
            AddLog(LogLevel.Error, $"Connection failed: {result.Error?.Message}");

            if (result.Error != null)
            {
                ShowErrorDialog("Connection Failed", result.Error);
            }
            else
            {
                ShowErrorDialog("Connection Failed", "Unable to connect to Exchange Online. Please check the logs for more details.");
            }
        }
    }

    private async Task DisconnectExchangeAsync(CancellationToken cancellationToken)
    {
        AddLog(LogLevel.Information, "Disconnecting from Exchange Online...");

        var result = await _workerService.DisconnectExchangeAsync(cancellationToken);

        if (result.IsSuccess)
        {
            ExchangeState = ConnectionState.Disconnected;
            ConnectedUser = null;
            ConnectedOrganization = null;
            Capabilities = null;
            IsGraphConnected = false;
            AddLog(LogLevel.Information, "Disconnected from Exchange Online");
        }
        else
        {
            AddLog(LogLevel.Error, $"Disconnect failed: {result.Error?.Message}");
        }
    }

    #endregion

    #region Event Handlers

    private void OnWorkerStateChanged(object? sender, WorkerConnectionState state)
    {
        RunOnUiThread(() =>
        {
            WorkerState = state;
            AddLog(LogLevel.Information, $"Worker state changed: {state}");

            if (state != WorkerConnectionState.Connected && ExchangeState == ConnectionState.Connected)
            {
                ExchangeState = ConnectionState.Disconnected;
                ConnectedUser = null;
                ConnectedOrganization = null;
                Capabilities = null;
                IsGraphConnected = false;
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
                        AddLog(logPayload.Level, logPayload.Message, logPayload.Source);
                    }
                    break;

                case EventType.Progress:
                    var progressPayload = JsonMessageSerializer.ExtractPayload<ProgressEventPayload>(evt.Payload);
                    if (progressPayload != null)
                    {
                        GlobalProgress = progressPayload.PercentComplete;
                        GlobalStatus = progressPayload.StatusMessage;
                        IsGlobalOperationRunning = progressPayload.PercentComplete < 100;
                    }
                    break;
            }
        });
    }

    private void OnCapabilitiesUpdated(object? sender, CapabilityMapDto capabilities)
    {
        RunOnUiThread(() =>
        {
            Capabilities = capabilities;
            var availableCount = capabilities.Cmdlets.Count(c => c.Value.IsAvailable);
            AddLog(LogLevel.Information, $"Capabilities updated: {availableCount} cmdlets available");
        });
    }

    private void OnNavigating(object? sender, NavigatingEventArgs e)
    {
        if (IsNavigationLocked)
        {
            RunOnUiThread(() =>
            {
                System.Windows.MessageBox.Show(
                    "An operation is currently running. Please wait for data loading or changes to finish before navigating.",
                    "Operation in Progress",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Information);
            });
            e.Cancel = true;
            return;
        }

        if (MailboxDetails != null && MailboxDetails.HasPendingChanges)
        {
            RunOnUiThread(() =>
            {
                var result = System.Windows.MessageBox.Show(
                    "You have unsaved changes in the Permission Manager. Do you want to discard them and continue?",
                    "Unsaved Changes",
                    System.Windows.MessageBoxButton.YesNo,
                    System.Windows.MessageBoxImage.Warning);

                if (result == System.Windows.MessageBoxResult.No)
                {
                    e.Cancel = true;
                }
            });
        }
    }

    private void OnPageChanged(object? sender, NavigationPage page)
    {
        RunOnUiThread(() => CurrentPage = page);
    }

    #endregion

    #region Helpers

    // Batched property change notifications for performance
    private static readonly string[] WorkerStateProperties =
    {
        nameof(WorkerStateDisplay),
        nameof(WorkerStateColor),
        nameof(CanStartWorker),
        nameof(CanStopWorker),
        nameof(CanRestartWorker),
        nameof(CanKillWorker),
        nameof(CanConnectExchange),
        nameof(CanDisconnectExchange),
        nameof(IsWorkerRunning)
    };

    private static readonly string[] ExchangeStateProperties =
    {
        nameof(ExchangeStateDisplay),
        nameof(ExchangeStateColor),
        nameof(IsExchangeConnected),
        nameof(CanConnectExchange),
        nameof(CanDisconnectExchange)
    };

    private static readonly string[] NavigationProperties =
    {
        nameof(IsDashboardPage),
        nameof(IsMailboxesPage),
        nameof(IsDeletedMailboxesPage),
        nameof(IsSharedMailboxesPage),
        nameof(IsMailboxSpacePage),
        nameof(IsDistributionListsPage),
        nameof(IsMessageTracePage),
        nameof(IsMailFlowPage),
        nameof(IsToolsPage),
        nameof(IsLogsPage),
        nameof(CurrentPageTitle)
    };

    private void NotifyWorkerStatePropertiesChanged()
    {
        foreach (var prop in WorkerStateProperties)
        {
            OnPropertyChanged(prop);
        }
        InvalidateCommandsOnUiThread();
    }

    private void NotifyExchangeStatePropertiesChanged()
    {
        foreach (var prop in ExchangeStateProperties)
        {
            OnPropertyChanged(prop);
        }
        InvalidateCommandsOnUiThread();
    }

    private void NotifyNavigationPropertiesChanged()
    {
        foreach (var prop in NavigationProperties)
        {
            OnPropertyChanged(prop);
        }
    }

    private void InvalidateCommandsOnUiThread()
    {
        var dispatcher = System.Windows.Application.Current?.Dispatcher;
        if (dispatcher == null)
        {
            CommandManager.InvalidateRequerySuggested();
            return;
        }

        if (dispatcher.CheckAccess())
        {
            CommandManager.InvalidateRequerySuggested();
        }
        else
        {
            dispatcher.BeginInvoke(DispatcherPriority.Background,
                new Action(() => CommandManager.InvalidateRequerySuggested()));
        }
    }

    public void AddLog(LogLevel level, string message, string? source = null)
    {
        if (!_isVerboseLoggingEnabled && level == LogLevel.Verbose)
        {
            return;
        }

        var entry = new LogEntry
        {
            Level = level,
            Message = message,
            Source = source ?? "UI"
        };

        LogEntries.Add(entry);

        if (LogEntries.Count > MaxLogEntries)
        {
            LogEntries.RemoveAt(0);
        }
    }

    private static bool IsEnvironmentFlagEnabled(string name)
    {
        var value = Environment.GetEnvironmentVariable(name);
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        return value.Equals("1", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("true", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("yes", StringComparison.OrdinalIgnoreCase) ||
               value.Equals("on", StringComparison.OrdinalIgnoreCase);
    }

    public void RegisterNavigationStateSource(INotifyPropertyChanged source)
    {
        if (_navigationStateSources.Contains(source))
        {
            return;
        }

        _navigationStateSources.Add(source);
        source.PropertyChanged += OnNavigationStateSourceChanged;
        UpdateNavigationLock();
    }

    public Task StartWorkerOnStartupAsync()
    {
        if (!CanStartWorker)
        {
            return Task.CompletedTask;
        }

        return StartWorkerAsync(CancellationToken.None);
    }

    private void OnNavigationStateSourceChanged(object? sender, PropertyChangedEventArgs e)
    {
        UpdateNavigationLock();
    }

    private void UpdateNavigationLock()
    {
        var isBlocked =
            IsGlobalOperationRunning ||
            (Dashboard?.IsLoading ?? false) ||
            (Mailboxes?.IsLoading ?? false) ||
            (DeletedMailboxes?.IsLoading ?? false) ||
            (SharedMailboxes?.IsLoading ?? false) ||
            (MailboxSpace?.IsLoading ?? false) ||
            (DistributionLists?.IsLoading ?? false) ||
            (DistributionLists?.IsLoadingDetails ?? false) ||
            (DistributionLists?.IsLoadingMembers ?? false) ||
            (MailboxDetails?.IsLoading ?? false) ||
            (MailboxDetails?.IsRetentionPolicyLoading ?? false) ||
            (MailboxDetails?.IsSaving ?? false);

        IsNavigationLocked = isBlocked;
    }

    public bool IsFeatureAvailable(Func<FeatureCapabilitiesDto, bool> featureCheck)
    {
        if (Capabilities?.Features == null) return false;
        return featureCheck(Capabilities.Features);
    }

    public string GetUnavailableTooltip(string featureName)
    {
        return $"{featureName} is not available with your current permissions";
    }

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

    #endregion

    public void Dispose()
    {
        _workerService.StateChanged -= OnWorkerStateChanged;
        _workerService.EventReceived -= OnEventReceived;
        _workerService.CapabilitiesUpdated -= OnCapabilitiesUpdated;
        _navigationService.PageChanged -= OnPageChanged;
        _navigationService.Navigating -= OnNavigating;

        foreach (var source in _navigationStateSources)
        {
            source.PropertyChanged -= OnNavigationStateSourceChanged;
        }
    }
}
