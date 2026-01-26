using System.Collections.ObjectModel;
using System.Windows.Input;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Application.UseCases;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Infrastructure.Ipc;
using ExchangeAdmin.Presentation.Helpers;

namespace ExchangeAdmin.Presentation.ViewModels;

             
                                     
              
public sealed class MainViewModel : ViewModelBase, IDisposable
{
    private readonly IWorkerService _workerService;
    private readonly ConnectExchangeUseCase _connectUseCase;
    private readonly DemoOperationUseCase _demoUseCase;

    private CancellationTokenSource? _demoOperationCts;

                   
    private WorkerConnectionState _workerState = WorkerConnectionState.NotStarted;
    private bool _isWorkerBusy;

                                
    private ConnectionState _exchangeState = ConnectionState.Disconnected;
    private string? _connectedUser;
    private string? _connectedOrganization;

                           
    private bool _isDemoRunning;
    private int _demoProgress;
    private string? _demoStatus;
    private int _demoDuration = 10;
    private int _demoItemCount = 10;
    private bool _demoSimulateError;

           
    private const int MaxLogEntries = 500;

    public MainViewModel(IWorkerService workerService)
    {
        _workerService = workerService;
        _connectUseCase = new ConnectExchangeUseCase(workerService);
        _demoUseCase = new DemoOperationUseCase(workerService);

                                
        _workerService.StateChanged += OnWorkerStateChanged;
        _workerService.EventReceived += OnEventReceived;

                              
        StartWorkerCommand = new AsyncRelayCommand(StartWorkerAsync, () => CanStartWorker);
        StopWorkerCommand = new AsyncRelayCommand(StopWorkerAsync, () => CanStopWorker);
        RestartWorkerCommand = new AsyncRelayCommand(RestartWorkerAsync, () => CanRestartWorker);
        KillWorkerCommand = new RelayCommand(() => _workerService.KillWorker(), () => CanKillWorker);

        ConnectExchangeCommand = new AsyncRelayCommand(ConnectExchangeAsync, () => CanConnectExchange);
        DisconnectExchangeCommand = new AsyncRelayCommand(DisconnectExchangeAsync, () => CanDisconnectExchange);

        RunDemoCommand = new AsyncRelayCommand(RunDemoAsync, () => CanRunDemo);
        CancelDemoCommand = new RelayCommand(CancelDemo, () => IsDemoRunning);

        ClearLogsCommand = new RelayCommand(() => LogEntries.Clear());
    }

    #region Properties

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
                OnPropertyChanged(nameof(CanRunDemo));
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
        ConnectionState.Connected => $"Connected: {ConnectedUser}",
        ConnectionState.Connecting => "Connecting...",
        ConnectionState.Reconnecting => "Reconnecting...",
        ConnectionState.Failed => "Connection Failed",
        _ => "Disconnected"
    };

    public string ExchangeStateColor => ExchangeState switch
    {
        ConnectionState.Connected => "#4EC9B0",
        ConnectionState.Connecting or ConnectionState.Reconnecting => "#DCDCAA",
        ConnectionState.Failed => "#F14C4C",
        _ => "#9D9D9D"
    };

    public bool IsExchangeConnected => ExchangeState == ConnectionState.Connected;

    public string? ConnectedUser
    {
        get => _connectedUser;
        private set => SetProperty(ref _connectedUser, value);
    }

    public string? ConnectedOrganization
    {
        get => _connectedOrganization;
        private set => SetProperty(ref _connectedOrganization, value);
    }

    public bool IsDemoRunning
    {
        get => _isDemoRunning;
        private set
        {
            if (SetProperty(ref _isDemoRunning, value))
            {
                OnPropertyChanged(nameof(CanRunDemo));
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }

    public int DemoProgress
    {
        get => _demoProgress;
        private set => SetProperty(ref _demoProgress, value);
    }

    public string? DemoStatus
    {
        get => _demoStatus;
        private set => SetProperty(ref _demoStatus, value);
    }

    public int DemoDuration
    {
        get => _demoDuration;
        set => SetProperty(ref _demoDuration, value);
    }

    public int DemoItemCount
    {
        get => _demoItemCount;
        set => SetProperty(ref _demoItemCount, value);
    }

    public bool DemoSimulateError
    {
        get => _demoSimulateError;
        set => SetProperty(ref _demoSimulateError, value);
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
                                       ExchangeState == ConnectionState.Disconnected;

    public bool CanDisconnectExchange => WorkerState == WorkerConnectionState.Connected &&
                                          ExchangeState == ConnectionState.Connected;

    public bool CanRunDemo => WorkerState == WorkerConnectionState.Connected && !IsDemoRunning;

    #endregion

    #region Commands

    public ICommand StartWorkerCommand { get; }
    public ICommand StopWorkerCommand { get; }
    public ICommand RestartWorkerCommand { get; }
    public ICommand KillWorkerCommand { get; }
    public ICommand ConnectExchangeCommand { get; }
    public ICommand DisconnectExchangeCommand { get; }
    public ICommand RunDemoCommand { get; }
    public ICommand CancelDemoCommand { get; }
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
                AddLog(LogLevel.Warning, "ExchangeOnlineManagement module is not available. Install it with: Install-Module ExchangeOnlineManagement");
            }
        }
        else
        {
            AddLog(LogLevel.Error, $"Failed to start worker: {_workerService.Status.LastError}");
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
            AddLog(LogLevel.Information, "Disconnected from Exchange Online");
        }
        else
        {
            AddLog(LogLevel.Error, $"Disconnect failed: {result.Error?.Message}");
        }
    }

    private async Task RunDemoAsync(CancellationToken cancellationToken)
    {
        IsDemoRunning = true;
        DemoProgress = 0;
        DemoStatus = "Starting...";

        _demoOperationCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);

        AddLog(LogLevel.Information, $"Starting demo operation: {DemoItemCount} items, {DemoDuration}s duration");

        var result = await _demoUseCase.ExecuteAsync(
            durationSeconds: DemoDuration,
            itemCount: DemoItemCount,
            simulateError: DemoSimulateError,
            errorAtPercent: 50,
            onLog: (level, msg) => RunOnUiThread(() => AddLog(level, msg)),
            onProgress: (percent, status) => RunOnUiThread(() =>
            {
                DemoProgress = percent;
                DemoStatus = status;
            }),
            onPartialOutput: item => RunOnUiThread(() =>
            {
                AddLog(LogLevel.Verbose, $"Partial output: Item {item.ItemId} - {item.Status}");
            }),
            cancellationToken: _demoOperationCts.Token);

        IsDemoRunning = false;
        _demoOperationCts?.Dispose();
        _demoOperationCts = null;

        if (result.IsSuccess && result.Value != null)
        {
            DemoStatus = $"Completed: {result.Value.ProcessedItems} items in {result.Value.ElapsedSeconds:F1}s";
            AddLog(LogLevel.Information, $"Demo completed: {result.Value.ProcessedItems} items processed");
        }
        else if (result.WasCancelled)
        {
            DemoStatus = "Cancelled";
            AddLog(LogLevel.Warning, "Demo operation cancelled");
        }
        else
        {
            DemoStatus = $"Error: {result.Error?.Message}";
            AddLog(LogLevel.Error, $"Demo operation failed: {result.Error?.Message}");
        }
    }

    private void CancelDemo()
    {
        if (_demoOperationCts != null)
        {
            AddLog(LogLevel.Information, "Cancelling demo operation...");
            _demoOperationCts.Cancel();
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
            }
        });
    }

    private void OnEventReceived(object? sender, EventEnvelope evt)
    {
        RunOnUiThread(() =>
        {
            if (evt.EventType == EventType.Log)
            {
                var payload = JsonMessageSerializer.ExtractPayload<LogEventPayload>(evt.Payload);
                if (payload != null)
                {
                    AddLog(payload.Level, payload.Message, payload.Source);
                }
            }
        });
    }

    #endregion

    #region Helpers

    private void AddLog(LogLevel level, string message, string? source = null)
    {
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

    #endregion

    public void Dispose()
    {
        _workerService.StateChanged -= OnWorkerStateChanged;
        _workerService.EventReceived -= OnEventReceived;
        _demoOperationCts?.Dispose();
    }
}
