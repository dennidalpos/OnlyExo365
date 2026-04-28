using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using OnlyExo365.Shell.Security;
using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts;
using OnlyExo365.Shell.UseCases;
using OnlyExo365.Contracts.Diagnostics;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Errors;
using OnlyExo365.Shell.Ipc;
using OnlyExo365.Shell.Helpers;
using OnlyExo365.Shell.Localization;

namespace OnlyExo365.Shell.ViewModels;

public sealed class ShellConnectionStateViewModel : ViewModelBase
{
    private const string DisableExchangeEnvVar = "ONLYEXO365_DISABLE_EXO";

    private readonly IConnectionWorkerService _workerService;
    private readonly ConnectExchangeUseCase _connectUseCase;
    private readonly PrepareInteractiveExchangeSignInUseCase? _prepareInteractiveSignInUseCase;
    private readonly ShellLogViewModel _logState;
    private readonly LeastPrivilegeEvaluator _leastPrivilegeEvaluator;
    private readonly ExchangeOnlineConfiguration _exchangeConfiguration;

    private WorkerConnectionState _workerState = WorkerConnectionState.NotStarted;
    private bool _isWorkerBusy;
    private ConnectionState _exchangeState = ConnectionState.Disconnected;
    private string? _connectedUser;
    private string? _connectedOrganization;
    private bool _isGraphConnected;
    private CapabilityMapDto? _capabilities;
    private bool _isWorkerConsoleVisible;
    private bool _isWorkerConsoleToggleBusy;
    private string? _workerStartupAlertDetails;

    public ShellConnectionStateViewModel(
        IConnectionWorkerService workerService,
        ShellLogViewModel logState,
        ExchangeOnlineConfiguration? exchangeConfiguration = null,
        IInteractiveExchangeBootstrapService? interactiveExchangeBootstrapService = null)
    {
        _workerService = workerService;
        _logState = logState;
        _connectUseCase = new ConnectExchangeUseCase(workerService);
        _exchangeConfiguration = exchangeConfiguration ?? ExchangeOnlineConfiguration.CreateDefault();
        _leastPrivilegeEvaluator = new LeastPrivilegeEvaluator(_exchangeConfiguration);
        if (interactiveExchangeBootstrapService != null)
        {
            _prepareInteractiveSignInUseCase = new PrepareInteractiveExchangeSignInUseCase(interactiveExchangeBootstrapService);
        }
        IsExchangeConnectionDisabled = IsEnvironmentFlagEnabled(DisableExchangeEnvVar);

        StartWorkerCommand = new AsyncRelayCommand(
            cancellationToken => StartWorkerAsync(cancellationToken, showFailureDialog: true),
            () => CanStartWorker);
        StopWorkerCommand = new AsyncRelayCommand(StopWorkerAsync, () => CanStopWorker);
        RestartWorkerCommand = new AsyncRelayCommand(RestartWorkerAsync, () => CanRestartWorker);
        KillWorkerCommand = new RelayCommand(() => _workerService.KillWorker(), () => CanKillWorker);
        SetWorkerConsoleVisibilityCommand = new AsyncRelayCommand<bool?>(
            SetWorkerConsoleVisibilityAsync,
            requestedVisibility => CanToggleWorkerConsole && requestedVisibility.HasValue);
        ConnectExchangeCommand = new AsyncRelayCommand(ConnectExchangeAsync, () => CanConnectExchange);
        DisconnectExchangeCommand = new AsyncRelayCommand(DisconnectExchangeAsync, () => CanDisconnectExchange);
    }

    public WorkerConnectionState WorkerState
    {
        get => _workerState;
        internal set
        {
            if (SetProperty(ref _workerState, value))
            {
                NotifyWorkerStatePropertiesChanged();
            }
        }
    }

    public string WorkerStateDisplay => WorkerState switch
    {
        WorkerConnectionState.NotStarted => Loc.Get("Status.Worker.NotStarted"),
        WorkerConnectionState.Starting => Loc.Get("Status.Worker.Starting"),
        WorkerConnectionState.WaitingForHandshake => Loc.Get("Status.Worker.Initializing"),
        WorkerConnectionState.Connected => Loc.Get("Status.Worker.Running"),
        WorkerConnectionState.Restarting => Loc.Get("Status.Worker.Restarting"),
        WorkerConnectionState.Stopped => Loc.Get("Status.Worker.Stopped"),
        WorkerConnectionState.Crashed => Loc.Get("Status.Worker.Crashed"),
        WorkerConnectionState.Unresponsive => Loc.Get("Status.Worker.Unresponsive"),
        _ => Loc.Get("Status.Unknown")
    };

    public string WorkerStateColor => WorkerState switch
    {
        WorkerConnectionState.Connected => "#4EC9B0",
        WorkerConnectionState.Starting or WorkerConnectionState.WaitingForHandshake or WorkerConnectionState.Restarting => "#DCDCAA",
        WorkerConnectionState.Crashed or WorkerConnectionState.Unresponsive => "#F14C4C",
        _ => "#9D9D9D"
    };

    public bool IsWorkerRunning => WorkerState == WorkerConnectionState.Connected;

    public string WorkerRunningDisplay => IsWorkerRunning
        ? Loc.Get("Common.Yes")
        : Loc.Get("Common.No");

    internal bool HasWorkerStartupAlert => !string.IsNullOrWhiteSpace(_workerStartupAlertDetails);

    internal string WorkerStartupAlertTitle => Loc.Get("Alert.WorkerStartupFailedTitle");

    internal string WorkerStartupAlertMessage => Loc.Get("Alert.WorkerStartupFailedMessage");

    internal string? WorkerStartupAlertDetails => _workerStartupAlertDetails;

    public bool IsWorkerConsoleVisible
    {
        get => _isWorkerConsoleVisible;
        internal set
        {
            if (SetProperty(ref _isWorkerConsoleVisible, value))
            {
                OnPropertyChanged(nameof(WorkerConsoleVisibilityDisplay));
            }
        }
    }

    public string WorkerConsoleVisibilityDisplay => IsWorkerConsoleVisible
        ? Loc.Get("Tools.WorkerConsoleVisible")
        : Loc.Get("Tools.WorkerConsoleHidden");

    public bool IsWorkerConsoleToggleBusy
    {
        get => _isWorkerConsoleToggleBusy;
        private set
        {
            if (SetProperty(ref _isWorkerConsoleToggleBusy, value))
            {
                OnPropertyChanged(nameof(CanToggleWorkerConsole));
                InvalidateCommandsOnUiThread();
            }
        }
    }

    public bool CanToggleWorkerConsole => IsWorkerRunning && !IsWorkerConsoleToggleBusy;

    public bool IsWorkerBusy
    {
        get => _isWorkerBusy;
        internal set => SetProperty(ref _isWorkerBusy, value);
    }

    public ConnectionState ExchangeState
    {
        get => _exchangeState;
        internal set
        {
            if (SetProperty(ref _exchangeState, value))
            {
                NotifyExchangeStatePropertiesChanged();
            }
        }
    }

    public string ExchangeStateDisplay => ExchangeState switch
    {
        _ when IsExchangeConnectionDisabled => Loc.Get("Status.Connection.DisabledPolicy"),
        ConnectionState.Connected => Loc.Get("Status.Connection.Connected"),
        ConnectionState.Connecting => Loc.Get("Status.Connection.Connecting"),
        ConnectionState.Reconnecting => Loc.Get("Status.Connection.Reconnecting"),
        ConnectionState.Failed => Loc.Get("Status.Connection.Failed"),
        _ => Loc.Get("Status.Connection.Disconnected")
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

    public bool IsExchangeConnectionDisabled { get; }

    public bool IsGraphConnected
    {
        get => _isGraphConnected;
        internal set
        {
            if (SetProperty(ref _isGraphConnected, value))
            {
                OnPropertyChanged(nameof(GraphStateColor));
                OnPropertyChanged(nameof(GraphStateDisplay));
                OnPropertyChanged(nameof(GraphStateTooltip));
            }
        }
    }

    public string GraphStateColor => IsGraphConnected
        ? "#4EC9B0"
        : IsExchangeConnected
            ? "#DCDCAA"
            : "#9D9D9D";

    public string GraphStateDisplay => IsGraphConnected
        ? Loc.Get("Status.Connection.Connected")
        : IsExchangeConnectionDisabled
            ? Loc.Get("Status.Connection.Disabled")
            : IsExchangeConnected
                ? Loc.Get("Status.Connection.Disconnected")
                : Loc.Get("Status.Connection.Disconnected");

    public string GraphStateTooltip => IsGraphConnected
        ? Loc.Get("Status.Graph.ActiveTooltip")
        : IsExchangeConnectionDisabled
            ? Loc.Get("Status.Graph.DisabledPolicyTooltip")
            : IsExchangeConnected
                ? Loc.Get("Status.Graph.RetryTooltip")
                : Loc.Get("Status.Graph.NotConnectedTooltip");

    public string? ConnectedUser
    {
        get => _connectedUser;
        internal set
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
        internal set => SetProperty(ref _connectedOrganization, value);
    }

    public CapabilityMapDto? Capabilities
    {
        get => _capabilities;
        internal set
        {
            if (SetProperty(ref _capabilities, value))
            {
                OnPropertyChanged(nameof(HasCapabilities));
                OnPropertyChanged(nameof(CanAccessMobileDevicesPage));
                OnPropertyChanged(nameof(MobileDevicesNavigationTooltip));
                OnPropertyChanged(nameof(CanAccessMigrationPage));
                OnPropertyChanged(nameof(MigrationNavigationTooltip));
                OnPropertyChanged(nameof(CanAccessPermissionsPage));
                OnPropertyChanged(nameof(PermissionsNavigationTooltip));
                OnPropertyChanged(nameof(CanAccessMessageTracePage));
                OnPropertyChanged(nameof(MessageTraceNavigationTooltip));
                OnPropertyChanged(nameof(CanAccessCompliancePage));
                OnPropertyChanged(nameof(ComplianceNavigationTooltip));
                OnPropertyChanged(nameof(CanAccessMailSecurityPage));
                OnPropertyChanged(nameof(MailSecurityNavigationTooltip));
                OnPropertyChanged(nameof(CapabilitiesDisplay));
            }
        }
    }

    public bool HasCapabilities => Capabilities != null;

    public string CapabilitiesDisplay => Capabilities == null
        ? Loc.Get("Tools.Capabilities.NotDetected")
        : Loc.GetFormat("Tools.Capabilities.CmdletsDetected", Capabilities.Cmdlets.Count);

    public bool CanAccessMobileDevicesPage => !IsExchangeConnected ||
                                              Capabilities == null ||
                                              MobileDevicesCapabilityState.From(Capabilities).IsModuleAvailable;

    public string MobileDevicesNavigationTooltip => CanAccessMobileDevicesPage
        ? Loc.Get("Nav.Tooltip.MobileDevices.Open")
        : MobileDevicesCapabilityState.From(Capabilities).Message ?? Loc.Get("Nav.Tooltip.MobileDevices.Unavailable");

    public bool CanAccessMigrationPage => CanAccessFeature(LeastPrivilegeCatalog.MigrationBatches);

    public string MigrationNavigationTooltip => GetFeatureNavigationTooltip(
        LeastPrivilegeCatalog.MigrationBatches,
        Loc.Get("Nav.Tooltip.Migration.Open"));

    public bool CanAccessPermissionsPage => CanAccessFeature(LeastPrivilegeCatalog.PermissionsRoleGroups);

    public string PermissionsNavigationTooltip => GetFeatureNavigationTooltip(
        LeastPrivilegeCatalog.PermissionsRoleGroups,
        Loc.Get("Nav.Tooltip.RoleGroups.Open"));

    public bool CanAccessMessageTracePage => CanAccessFeature(LeastPrivilegeCatalog.MessageTraceRead);

    public string MessageTraceNavigationTooltip => GetFeatureNavigationTooltip(
        LeastPrivilegeCatalog.MessageTraceRead,
        Loc.Get("Nav.Tooltip.MessageTrace.Open"));

    public bool CanAccessCompliancePage => CanAccessFeature(LeastPrivilegeCatalog.ComplianceAuditAndEDiscovery);

    public string ComplianceNavigationTooltip => GetFeatureNavigationTooltip(
        LeastPrivilegeCatalog.ComplianceAuditAndEDiscovery,
        Loc.Get("Nav.Tooltip.Compliance.Open"));

    public bool CanAccessMailSecurityPage => CanAccessFeature(LeastPrivilegeCatalog.MailSecurityBaseline);

    public string MailSecurityNavigationTooltip => GetFeatureNavigationTooltip(
        LeastPrivilegeCatalog.MailSecurityBaseline,
        Loc.Get("Nav.Tooltip.MailSecurity.Open"));

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

    public ICommand StartWorkerCommand { get; }
    public ICommand StopWorkerCommand { get; }
    public ICommand RestartWorkerCommand { get; }
    public ICommand KillWorkerCommand { get; }
    public ICommand SetWorkerConsoleVisibilityCommand { get; }
    public ICommand ConnectExchangeCommand { get; }
    public ICommand DisconnectExchangeCommand { get; }

    public void ApplyWorkerStateChange(WorkerConnectionState state)
    {
        WorkerState = state;
        IsWorkerConsoleVisible = state == WorkerConnectionState.Connected && _workerService.Status.IsConsoleVisible;
        _logState.AddLog(LogLevel.Information, $"Worker state changed: {state}");

        if (state == WorkerConnectionState.Connected)
        {
            ClearWorkerStartupAlert();
        }

        if (state != WorkerConnectionState.Connected && ExchangeState == ConnectionState.Connected)
        {
            ResetExchangeSession();
        }
    }

    public void ApplyCapabilities(CapabilityMapDto capabilities)
    {
        Capabilities = capabilities;
        var availableCount = capabilities.Cmdlets.Count(cmdlet => cmdlet.Value.IsAvailable);
        _logState.AddLog(LogLevel.Information, $"Capabilities updated: {availableCount} cmdlets available");
    }

    public void ResetExchangeSession()
    {
        ExchangeState = ConnectionState.Disconnected;
        ConnectedUser = null;
        ConnectedOrganization = null;
        Capabilities = null;
        IsGraphConnected = false;
        if (!IsWorkerRunning)
        {
            IsWorkerConsoleVisible = false;
        }
    }

    public async Task RefreshConnectionStatusAsync(CancellationToken cancellationToken = default)
    {
        if (WorkerState != WorkerConnectionState.Connected)
        {
            return;
        }

        try
        {
            var result = await _workerService.GetConnectionStatusAsync(cancellationToken);
            if (!result.IsSuccess || result.Value == null)
            {
                return;
            }

            await RunOnUiThreadAsync(() =>
            {
                ExchangeState = result.Value.State;
                ConnectedUser = result.Value.UserPrincipalName;
                ConnectedOrganization = result.Value.Organization;
                IsGraphConnected = result.Value.GraphConnected;

                if (result.Value.State != ConnectionState.Connected)
                {
                    Capabilities = null;
                }
            });
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            _logState.AddLog(LogLevel.Warning, $"Unable to refresh connection status: {ex.Message}");
        }
    }

    public Task StartWorkerOnStartupAsync()
    {
        if (!CanStartWorker)
        {
            return Task.CompletedTask;
        }

        return StartWorkerAsync(CancellationToken.None, showFailureDialog: false);
    }

    public bool IsFeatureAvailable(Func<FeatureCapabilitiesDto, bool> featureCheck)
    {
        if (Capabilities?.Features == null)
        {
            return false;
        }

        return featureCheck(Capabilities.Features);
    }

    public string GetUnavailableTooltip(string featureName)
    {
        return $"{featureName} is not available with your current permissions";
    }

    public LeastPrivilegeFeatureEvaluation EvaluateLeastPrivilege(string featureId)
        => _leastPrivilegeEvaluator.Evaluate(featureId, Capabilities);

    private async Task StartWorkerAsync(CancellationToken cancellationToken, bool showFailureDialog)
    {
        ClearWorkerStartupAlert();
        _logState.AddLog(LogLevel.Information, "Starting worker...");

        var success = await _workerService.StartWorkerAsync(cancellationToken);
        if (success)
        {
            ClearWorkerStartupAlert();
            _logState.AddLog(LogLevel.Information, "Worker started successfully");

            var status = _workerService.Status;
            if (!status.IsModuleAvailable)
            {
                var warning = Loc.Get("Warn.ModuleMissing.Message");
                _logState.AddLog(LogLevel.Warning, "ExchangeOnlineManagement module is not available. Install it with: Install-Module ExchangeOnlineManagement");
                ErrorDialogService.ShowWarning(Loc.Get("Warn.ModuleMissing.Title"), warning);
            }

            return;
        }

        var errorMessage = _workerService.Status.LastError;
        var errorDetails = BuildWorkerStartupErrorDetails(errorMessage);
        _logState.AddLog(LogLevel.Error, $"Failed to start worker: {errorMessage}");
        if (showFailureDialog)
        {
            ShowErrorDialog(
                Loc.Get("Error.WorkerStartup.Title"),
                Loc.Get("Error.WorkerStartup.Message"),
                errorDetails);
            return;
        }

        SetWorkerStartupAlert(errorDetails);
    }

    private async Task StopWorkerAsync(CancellationToken cancellationToken)
    {
        _logState.AddLog(LogLevel.Information, "Stopping worker...");
        await _workerService.StopWorkerAsync();
        _logState.AddLog(LogLevel.Information, "Worker stopped");
    }

    private async Task RestartWorkerAsync(CancellationToken cancellationToken)
    {
        _logState.AddLog(LogLevel.Information, "Restarting worker...");
        var success = await _workerService.RestartWorkerAsync(cancellationToken);

        if (success)
        {
            _logState.AddLog(LogLevel.Information, "Worker restarted successfully");
        }
        else
        {
            _logState.AddLog(LogLevel.Error, $"Failed to restart worker: {_workerService.Status.LastError}");
        }
    }

    private async Task SetWorkerConsoleVisibilityAsync(bool? requestedVisibility, CancellationToken cancellationToken)
    {
        if (!requestedVisibility.HasValue)
        {
            return;
        }

        var previousVisibility = IsWorkerConsoleVisible;
        IsWorkerConsoleToggleBusy = true;
        IsWorkerConsoleVisible = requestedVisibility.Value;

        try
        {
            var result = await _workerService.SetWorkerConsoleVisibilityAsync(
                requestedVisibility.Value,
                cancellationToken: cancellationToken);

            if (result.IsSuccess && result.Value != null)
            {
                IsWorkerConsoleVisible = result.Value.IsVisible;
                _logState.AddLog(
                    LogLevel.Information,
                    result.Value.IsVisible ? "Worker console shown" : "Worker console hidden");
                return;
            }

            IsWorkerConsoleVisible = previousVisibility;
            _logState.AddLog(
                LogLevel.Error,
                $"Failed to change worker console visibility: {result.Error?.Message ?? "Unknown error"}");
        }
        catch (OperationCanceledException)
        {
            IsWorkerConsoleVisible = previousVisibility;
        }
        catch (Exception ex)
        {
            IsWorkerConsoleVisible = previousVisibility;
            _logState.AddLog(LogLevel.Error, $"Failed to change worker console visibility: {ex.Message}");
        }
        finally
        {
            IsWorkerConsoleToggleBusy = false;
        }
    }

    private async Task ConnectExchangeAsync(CancellationToken cancellationToken)
    {
        if (IsExchangeConnectionDisabled)
        {
            _logState.AddLog(LogLevel.Warning, "Exchange Online connections are disabled by policy (ONLYEXO365_DISABLE_EXO=1).");
            ErrorDialogService.ShowWarning(
                Loc.Get("Warn.ConnectionDisabled.Title"),
                Loc.Get("Warn.ConnectionDisabled.Message"));
            return;
        }

        ExchangeState = ConnectionState.Connecting;
        _logState.AddLog(LogLevel.Information, "Connecting to Exchange Online...");

        if (_exchangeConfiguration.AuthenticationMode == ExchangeAuthenticationMode.Interactive &&
            _prepareInteractiveSignInUseCase != null)
        {
            var bootstrapResult = await _prepareInteractiveSignInUseCase.ExecuteAsync(
                onLog: (level, message) => RunOnUiThread(() => _logState.AddLog(level, message)),
                cancellationToken: cancellationToken);

            if (bootstrapResult.WasCancelled)
            {
                ResetExchangeSession();
                _logState.AddLog(LogLevel.Warning, "Interactive sign-in cancelled");
                return;
            }

            if (!bootstrapResult.IsSuccess)
            {
                ExchangeState = ConnectionState.Failed;
                IsGraphConnected = false;
                _logState.AddLog(LogLevel.Error, $"Interactive sign-in failed: {bootstrapResult.Error?.Message}");

                if (bootstrapResult.Error != null)
                {
                    ShowErrorDialog("Interactive Sign-In Failed", bootstrapResult.Error);
                }

                return;
            }
        }

        var result = await _connectUseCase.ExecuteAsync(
            onLog: (level, message) => RunOnUiThread(() => _logState.AddLog(level, message)),
            cancellationToken: cancellationToken);

        if (result.IsSuccess && result.Value != null)
        {
            ExchangeState = result.Value.State;
            ConnectedUser = result.Value.UserPrincipalName;
            ConnectedOrganization = result.Value.Organization;
            IsGraphConnected = result.Value.GraphConnected;
            _logState.AddLog(LogLevel.Information, $"Connected as {ConnectedUser} to {ConnectedOrganization}");
            return;
        }

        if (result.WasCancelled)
        {
            ResetExchangeSession();
            _logState.AddLog(LogLevel.Warning, "Connection cancelled");
            return;
        }

        ExchangeState = ConnectionState.Failed;
        IsGraphConnected = false;
        _logState.AddLog(LogLevel.Error, $"Connection failed: {result.Error?.Message}");

        if (result.Error != null)
        {
            ShowErrorDialog("Connection Failed", result.Error);
            return;
        }

        ShowErrorDialog("Connection Failed", "Unable to connect to Exchange Online. Please check the logs for more details.");
    }

    private async Task DisconnectExchangeAsync(CancellationToken cancellationToken)
    {
        _logState.AddLog(LogLevel.Information, "Disconnecting from Exchange Online...");

        var result = await _workerService.DisconnectExchangeAsync(cancellationToken);
        if (result.IsSuccess)
        {
            ResetExchangeSession();
            _logState.AddLog(LogLevel.Information, "Disconnected from Exchange Online");
        }
        else
        {
            _logState.AddLog(LogLevel.Error, $"Disconnect failed: {result.Error?.Message}");
        }
    }

    private void NotifyWorkerStatePropertiesChanged()
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
        OnPropertyChanged(nameof(WorkerRunningDisplay));
        OnPropertyChanged(nameof(CanToggleWorkerConsole));
        OnPropertyChanged(nameof(WorkerConsoleVisibilityDisplay));
        InvalidateCommandsOnUiThread();
    }

    private void NotifyExchangeStatePropertiesChanged()
    {
        OnPropertyChanged(nameof(ExchangeStateDisplay));
        OnPropertyChanged(nameof(ExchangeStateColor));
        OnPropertyChanged(nameof(IsExchangeConnected));
        OnPropertyChanged(nameof(CapabilitiesDisplay));
        OnPropertyChanged(nameof(GraphStateColor));
        OnPropertyChanged(nameof(GraphStateDisplay));
        OnPropertyChanged(nameof(GraphStateTooltip));
        OnPropertyChanged(nameof(CanAccessMobileDevicesPage));
        OnPropertyChanged(nameof(MobileDevicesNavigationTooltip));
        OnPropertyChanged(nameof(CanAccessMigrationPage));
        OnPropertyChanged(nameof(MigrationNavigationTooltip));
        OnPropertyChanged(nameof(CanAccessPermissionsPage));
        OnPropertyChanged(nameof(PermissionsNavigationTooltip));
        OnPropertyChanged(nameof(CanAccessMessageTracePage));
        OnPropertyChanged(nameof(MessageTraceNavigationTooltip));
        OnPropertyChanged(nameof(CanAccessCompliancePage));
        OnPropertyChanged(nameof(ComplianceNavigationTooltip));
        OnPropertyChanged(nameof(CanAccessMailSecurityPage));
        OnPropertyChanged(nameof(MailSecurityNavigationTooltip));
        OnPropertyChanged(nameof(CanConnectExchange));
        OnPropertyChanged(nameof(CanDisconnectExchange));
        InvalidateCommandsOnUiThread();
    }

    private bool CanAccessFeature(string featureId)
    {
        if (!IsExchangeConnected)
        {
            return true;
        }

        var evaluation = _leastPrivilegeEvaluator.Evaluate(featureId, Capabilities);
        return evaluation.IsNavigationAllowed;
    }

    private string GetFeatureNavigationTooltip(string featureId, string defaultValue)
    {
        if (!IsExchangeConnected)
        {
            return defaultValue;
        }

        var evaluation = _leastPrivilegeEvaluator.Evaluate(featureId, Capabilities);
        if (evaluation.IsNavigationAllowed)
        {
            return evaluation.Status == LeastPrivilegeFeatureStatus.Available
                ? defaultValue
                : $"{defaultValue}. {evaluation.ValidationMessage}";
        }

        return evaluation.MissingRequirementsDisplay;
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
            dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(CommandManager.InvalidateRequerySuggested));
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

    internal string BuildWorkerStartupErrorDetails(string? errorMessage)
    {
        var normalizedError = string.IsNullOrWhiteSpace(errorMessage)
            ? Loc.Get("Status.Unknown")
            : errorMessage.Trim();

        var guidance = IsHandshakeStartupError(normalizedError)
            ? Loc.Get("Error.WorkerStartup.Guidance.Handshake")
            : Loc.Get("Error.WorkerStartup.Guidance.Generic");

        return string.Join(
            Environment.NewLine + Environment.NewLine,
            Loc.GetFormat("Error.WorkerStartup.DetailsFormat", normalizedError),
            guidance);
    }

    internal static bool IsHandshakeStartupError(string errorMessage)
    {
        return errorMessage.Contains("handshake", StringComparison.OrdinalIgnoreCase) ||
               errorMessage.Contains("ipc session", StringComparison.OrdinalIgnoreCase) ||
               errorMessage.Contains("session validation", StringComparison.OrdinalIgnoreCase);
    }

    private void SetWorkerStartupAlert(string details)
    {
        var normalizedDetails = string.IsNullOrWhiteSpace(details)
            ? null
            : details.Trim();

        if (string.Equals(_workerStartupAlertDetails, normalizedDetails, StringComparison.Ordinal))
        {
            return;
        }

        _workerStartupAlertDetails = normalizedDetails;
        OnPropertyChanged(nameof(HasWorkerStartupAlert));
        OnPropertyChanged(nameof(WorkerStartupAlertTitle));
        OnPropertyChanged(nameof(WorkerStartupAlertMessage));
        OnPropertyChanged(nameof(WorkerStartupAlertDetails));
    }

    private void ClearWorkerStartupAlert()
    {
        if (string.IsNullOrWhiteSpace(_workerStartupAlertDetails))
        {
            return;
        }

        _workerStartupAlertDetails = null;
        OnPropertyChanged(nameof(HasWorkerStartupAlert));
        OnPropertyChanged(nameof(WorkerStartupAlertTitle));
        OnPropertyChanged(nameof(WorkerStartupAlertMessage));
        OnPropertyChanged(nameof(WorkerStartupAlertDetails));
    }

    private void ShowErrorDialog(string title, string message, string? details = null)
    {
        RunOnUiThread(() => ErrorDialogService.ShowError(title, message, details));
    }

    private void ShowErrorDialog(string title, NormalizedErrorDto error)
    {
        RunOnUiThread(() => ErrorDialogService.ShowError(title, error));
    }

    private void ShowErrorDialog(string title, NormalizedError error)
    {
        RunOnUiThread(() => ErrorDialogService.ShowError(title, error));
    }
}

