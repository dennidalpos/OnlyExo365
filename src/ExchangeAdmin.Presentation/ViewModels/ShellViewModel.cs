using System.Collections.ObjectModel;
using System.Windows.Input;
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

                 
    private NavigationPage _currentPage = NavigationPage.Dashboard;

                   
    private CapabilityMapDto? _capabilities;

                      
    private bool _isGlobalOperationRunning;
    private int _globalProgress;
    private string? _globalStatus;
    private readonly bool _isExchangeConnectionDisabled;

              
    private bool _isVerboseLoggingEnabled = false;
    private const int MaxLogEntries = 1000;

                       
    public DashboardViewModel? Dashboard { get; set; }
    public MailboxListViewModel? Mailboxes { get; set; }
    public MailboxListViewModel? SharedMailboxes { get; set; }
    public MailboxDetailsViewModel? MailboxDetails { get; set; }
    public MailboxSpaceViewModel? MailboxSpace { get; set; }
    public DistributionListViewModel? DistributionLists { get; set; }
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
        NavigateToSharedMailboxesCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.SharedMailboxes));
        NavigateToMailboxSpaceCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.MailboxSpace));
        NavigateToDistributionListsCommand = new RelayCommand(() => _navigationService.NavigateTo(NavigationPage.DistributionLists));
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
                OnPropertyChanged(nameof(WorkerStateDisplay));
                OnPropertyChanged(nameof(WorkerStateColor));
                OnPropertyChanged(nameof(CanStartWorker));
                OnPropertyChanged(nameof(CanStopWorker));
                OnPropertyChanged(nameof(CanRestartWorker));
                OnPropertyChanged(nameof(CanKillWorker));
                OnPropertyChanged(nameof(CanConnectExchange));
                OnPropertyChanged(nameof(CanDisconnectExchange));
                OnPropertyChanged(nameof(IsWorkerRunning));
                CommandManager.InvalidateRequerySuggested();
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
                OnPropertyChanged(nameof(ExchangeStateDisplay));
                OnPropertyChanged(nameof(ExchangeStateColor));
                OnPropertyChanged(nameof(IsExchangeConnected));
                OnPropertyChanged(nameof(CanConnectExchange));
                OnPropertyChanged(nameof(CanDisconnectExchange));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public string ExchangeStateDisplay => ExchangeState switch
    {
        _ when IsExchangeConnectionDisabled => "Disabled (policy)",
        ConnectionState.Connected => $"Connected: {ConnectedUser}",
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
                OnPropertyChanged(nameof(IsDashboardPage));
                OnPropertyChanged(nameof(IsMailboxesPage));
                OnPropertyChanged(nameof(IsSharedMailboxesPage));
                OnPropertyChanged(nameof(IsMailboxSpacePage));
                OnPropertyChanged(nameof(IsDistributionListsPage));
                OnPropertyChanged(nameof(IsToolsPage));
                OnPropertyChanged(nameof(IsLogsPage));
                OnPropertyChanged(nameof(CurrentPageTitle));
            }
        }
    }

    public bool IsDashboardPage => CurrentPage == NavigationPage.Dashboard;
    public bool IsMailboxesPage => CurrentPage == NavigationPage.Mailboxes;
    public bool IsSharedMailboxesPage => CurrentPage == NavigationPage.SharedMailboxes;
    public bool IsMailboxSpacePage => CurrentPage == NavigationPage.MailboxSpace;
    public bool IsDistributionListsPage => CurrentPage == NavigationPage.DistributionLists;
    public bool IsToolsPage => CurrentPage == NavigationPage.Tools;
    public bool IsLogsPage => CurrentPage == NavigationPage.Logs;

    public string CurrentPageTitle => CurrentPage switch
    {
        NavigationPage.Dashboard => "Dashboard",
        NavigationPage.Mailboxes => "Mailboxes",
        NavigationPage.SharedMailboxes => "Shared Mailboxes",
        NavigationPage.MailboxSpace => "Spazio mailbox",
        NavigationPage.DistributionLists => "Distribution Lists",
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
        private set => SetProperty(ref _isGlobalOperationRunning, value);
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
    public ICommand NavigateToSharedMailboxesCommand { get; }
    public ICommand NavigateToMailboxSpaceCommand { get; }
    public ICommand NavigateToDistributionListsCommand { get; }
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
            AddLog(LogLevel.Information, $"Connected as {ConnectedUser} to {ConnectedOrganization}");
        }
        else if (result.WasCancelled)
        {
            ExchangeState = ConnectionState.Disconnected;
            AddLog(LogLevel.Warning, "Connection cancelled");
        }
        else
        {
            ExchangeState = ConnectionState.Failed;
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

        LogEntries.Insert(0, entry);

                                
        while (LogEntries.Count > MaxLogEntries)
        {
            LogEntries.RemoveAt(LogEntries.Count - 1);
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
    }
}
