using OnlyExo365.Contracts.Diagnostics;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Errors;
using OnlyExo365.Contracts.Results;
using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts;
using OnlyExo365.Shell.Ipc;
using OnlyExo365.Shell.ViewModels;
using OnlyExo365.Shell.Localization;

namespace OnlyExo365.Tests;

public sealed class ShellConnectionStateViewModelTests : IDisposable
{
    private readonly string _tempDirectory = Path.Combine(Path.GetTempPath(), "OnlyExo365.Tests", Guid.NewGuid().ToString("N"));

    [WpfFact]
    public async Task SetWorkerConsoleVisibilityCommand_UpdatesStateAndLogsOnSuccess()
    {
        var workerService = new WorkerConsoleConnectionWorkerServiceStub();
        var viewModel = CreateViewModel(workerService, out var logState);

        viewModel.ApplyWorkerStateChange(WorkerConnectionState.Connected);
        workerService.SetToggleResult(Result<SetWorkerConsoleVisibilityResponse>.Success(new SetWorkerConsoleVisibilityResponse
        {
            IsVisible = true,
            Message = "Worker console shown."
        }));

        viewModel.SetWorkerConsoleVisibilityCommand.Execute(true);
        await WaitForAsync(() => workerService.LastRequestedVisibility.HasValue);

        Assert.True(viewModel.IsWorkerConsoleVisible);
        Assert.Equal(true, workerService.LastRequestedVisibility);
        Assert.Contains(logState.LogEntries, entry => entry.Message.Contains("Worker console shown", StringComparison.Ordinal));
    }

    [WpfFact]
    public async Task SetWorkerConsoleVisibilityCommand_RestoresPreviousStateOnFailure()
    {
        var workerService = new WorkerConsoleConnectionWorkerServiceStub(initialConsoleVisibility: false);
        var viewModel = CreateViewModel(workerService, out var logState);

        viewModel.ApplyWorkerStateChange(WorkerConnectionState.Connected);
        workerService.SetToggleResult(Result<SetWorkerConsoleVisibilityResponse>.Failure(
            NormalizedError.Create(ErrorCode.Unknown, "Boom")));

        viewModel.SetWorkerConsoleVisibilityCommand.Execute(true);
        await WaitForAsync(() => workerService.LastRequestedVisibility.HasValue);

        Assert.False(viewModel.IsWorkerConsoleVisible);
        Assert.Contains(logState.LogEntries, entry => entry.Message.Contains("Failed to change worker console visibility", StringComparison.Ordinal));
    }

    [WpfFact]
    public async Task SetWorkerConsoleVisibilityCommand_AppliesRequestedStateImmediately()
    {
        var workerService = new WorkerConsoleConnectionWorkerServiceStub(initialConsoleVisibility: false);
        var viewModel = CreateViewModel(workerService, out _);
        var toggleCompletion = new TaskCompletionSource<Result<SetWorkerConsoleVisibilityResponse>>(TaskCreationOptions.RunContinuationsAsynchronously);
        workerService.SetToggleTask(toggleCompletion.Task);

        viewModel.ApplyWorkerStateChange(WorkerConnectionState.Connected);
        viewModel.SetWorkerConsoleVisibilityCommand.Execute(true);
        await WaitForAsync(() => workerService.LastRequestedVisibility.HasValue);

        Assert.True(viewModel.IsWorkerConsoleVisible);
        Assert.True(viewModel.IsWorkerConsoleToggleBusy);

        toggleCompletion.SetResult(Result<SetWorkerConsoleVisibilityResponse>.Success(new SetWorkerConsoleVisibilityResponse
        {
            IsVisible = true,
            Message = "Worker console shown."
        }));
        await WaitForAsync(() => !viewModel.IsWorkerConsoleToggleBusy);
    }

    [WpfFact]
    public async Task ConnectExchange_InteractiveMode_PrimesInteractiveBootstrapBeforeWorkerConnect()
    {
        var workerService = new ConnectWorkerServiceStub();
        var bootstrapService = new InteractiveBootstrapServiceStub();
        var viewModel = CreateViewModel(
            workerService,
            out _,
            new ExchangeOnlineConfiguration { AuthenticationMode = ExchangeAuthenticationMode.Interactive },
            bootstrapService);

        viewModel.ApplyWorkerStateChange(WorkerConnectionState.Connected);
        viewModel.ConnectExchangeCommand.Execute(null);

        await WaitForAsync(() => workerService.ConnectInvoked);

        Assert.True(bootstrapService.WasInvoked);
        Assert.True(workerService.ConnectInvoked);
    }

    [WpfFact]
    public async Task ConnectExchange_DoesNotCallWorkerWhenInteractiveBootstrapFails()
    {
        var workerService = new ConnectWorkerServiceStub();
        var bootstrapService = new InteractiveBootstrapServiceStub
        {
            Result = Result.Failure(NormalizedError.Create(ErrorCode.AuthenticationFailed, "Popup failed"))
        };
        var viewModel = CreateViewModel(
            workerService,
            out var logState,
            new ExchangeOnlineConfiguration { AuthenticationMode = ExchangeAuthenticationMode.Interactive },
            bootstrapService);

        viewModel.ApplyWorkerStateChange(WorkerConnectionState.Connected);
        viewModel.ConnectExchangeCommand.Execute(null);

        await WaitForAsync(() => bootstrapService.WasInvoked);

        Assert.False(workerService.ConnectInvoked);
        Assert.Equal(ConnectionState.Failed, viewModel.ExchangeState);
        Assert.Contains(logState.LogEntries, entry => entry.Message.Contains("Interactive sign-in failed", StringComparison.Ordinal));
    }

    [WpfFact]
    public void BuildWorkerStartupErrorDetails_UsesHandshakeGuidanceForHandshakeFailures()
    {
        var previousLocale = LocalizationService.Instance.CurrentLocale;
        LocalizationService.Instance.SetLocale("en");

        try
        {
            var viewModel = CreateViewModel(new WorkerConsoleConnectionWorkerServiceStub(), out _);

            var details = viewModel.BuildWorkerStartupErrorDetails("Handshake failed: IPC session validation failed.");

            Assert.Contains("Error: Handshake failed: IPC session validation failed.", details, StringComparison.Ordinal);
            Assert.Contains("Close every running OnlyExo365 window and worker process.", details, StringComparison.Ordinal);
            Assert.DoesNotContain("Verify PowerShell 7 is installed.", details, StringComparison.Ordinal);
        }
        finally
        {
            LocalizationService.Instance.SetLocale(previousLocale);
        }
    }

    [WpfFact]
    public async Task StartWorkerOnStartupAsync_CapturesInlineStartupAlertDetails()
    {
        var previousLocale = LocalizationService.Instance.CurrentLocale;
        LocalizationService.Instance.SetLocale("en");

        try
        {
            var workerService = new WorkerConsoleConnectionWorkerServiceStub();
            workerService.SetStartWorkerResult(
                success: false,
                lastError: "Worker executable not found: OnlyExo365.Worker.exe",
                resultingState: WorkerConnectionState.Crashed);
            var viewModel = CreateViewModel(workerService, out _);

            await viewModel.StartWorkerOnStartupAsync();

            Assert.True(viewModel.HasWorkerStartupAlert);
            Assert.Equal("Worker startup failed", viewModel.WorkerStartupAlertTitle);
            Assert.Contains("Use Start Worker to retry without leaving the current page.", viewModel.WorkerStartupAlertMessage, StringComparison.Ordinal);
            Assert.Contains("Worker executable not found: OnlyExo365.Worker.exe", viewModel.WorkerStartupAlertDetails, StringComparison.Ordinal);
        }
        finally
        {
            LocalizationService.Instance.SetLocale(previousLocale);
        }
    }

    [WpfFact]
    public async Task StartWorkerCommand_WhenFailureOccurs_DoesNotReuseStartupBannerFeedback()
    {
        var workerService = new WorkerConsoleConnectionWorkerServiceStub();
        workerService.SetStartWorkerResult(
            success: false,
            lastError: "Worker executable not found: OnlyExo365.Worker.exe",
            resultingState: WorkerConnectionState.Crashed);
        var viewModel = CreateViewModel(workerService, out var logState);

        viewModel.StartWorkerCommand.Execute(null);
        await WaitForAsync(() => logState.LogEntries.Any(entry => entry.Message.Contains("Failed to start worker", StringComparison.Ordinal)));

        Assert.False(viewModel.HasWorkerStartupAlert);
    }

    [WpfFact]
    public void NavigationTooltips_FollowSelectedLocale()
    {
        var previousLocale = LocalizationService.Instance.CurrentLocale;
        LocalizationService.Instance.SetLocale("it");

        try
        {
            var viewModel = CreateViewModel(new WorkerConsoleConnectionWorkerServiceStub(), out _);

            Assert.Equal("Apri l'area di lavoro Migrazione", viewModel.MigrationNavigationTooltip);
            Assert.Equal("Apri l'area di lavoro Traccia messaggi", viewModel.MessageTraceNavigationTooltip);
        }
        finally
        {
            LocalizationService.Instance.SetLocale(previousLocale);
        }
    }

    [WpfFact]
    public void CanToggleWorkerConsole_IsFalseWhenWorkerIsStopped()
    {
        var viewModel = CreateViewModel(new WorkerConsoleConnectionWorkerServiceStub(), out _);

        viewModel.ApplyWorkerStateChange(WorkerConnectionState.Stopped);

        Assert.False(viewModel.CanToggleWorkerConsole);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDirectory))
            {
                Directory.Delete(_tempDirectory, recursive: true);
            }
        }
        catch
        {
        }
    }

    private ShellConnectionStateViewModel CreateViewModel(
        WorkerConsoleConnectionWorkerServiceStub workerService,
        out ShellLogViewModel logState,
        ExchangeOnlineConfiguration? configuration = null,
        IInteractiveExchangeBootstrapService? bootstrapService = null)
    {
        Directory.CreateDirectory(_tempDirectory);
        logState = new ShellLogViewModel(new PersistentLogWriter("ui-tests", _tempDirectory));
        return new ShellConnectionStateViewModel(workerService, logState, configuration, bootstrapService);
    }

    private static async Task WaitForAsync(Func<bool> condition, int timeoutMs = 1000)
    {
        var startedAt = Environment.TickCount64;
        while (!condition())
        {
            if (Environment.TickCount64 - startedAt > timeoutMs)
            {
                throw new TimeoutException("Condition not met in time.");
            }

            await Task.Delay(20);
        }
    }

    private class WorkerConsoleConnectionWorkerServiceStub : TestConnectionWorkerServiceBase
    {
        private WorkerStatus _status;
        private Result<SetWorkerConsoleVisibilityResponse> _toggleResult;
        private Task<Result<SetWorkerConsoleVisibilityResponse>>? _toggleTask;
        private bool _startWorkerResult = true;

        public WorkerConsoleConnectionWorkerServiceStub(bool initialConsoleVisibility = false)
        {
            _status = new WorkerStatus
            {
                State = WorkerConnectionState.Connected,
                IsModuleAvailable = true,
                IsConsoleVisible = initialConsoleVisibility
            };
            _toggleResult = Result<SetWorkerConsoleVisibilityResponse>.Success(new SetWorkerConsoleVisibilityResponse
            {
                IsVisible = initialConsoleVisibility,
                Message = initialConsoleVisibility ? "Worker console shown." : "Worker console hidden."
            });
        }

        public bool? LastRequestedVisibility { get; private set; }

        public override WorkerStatus Status => _status;

        public void SetToggleResult(Result<SetWorkerConsoleVisibilityResponse> result)
        {
            _toggleResult = result;
            _toggleTask = null;
        }

        public void SetToggleTask(Task<Result<SetWorkerConsoleVisibilityResponse>> task)
        {
            _toggleTask = task;
        }

        public void SetStartWorkerResult(bool success, string? lastError = null, WorkerConnectionState resultingState = WorkerConnectionState.Connected)
        {
            _startWorkerResult = success;
            _status = new WorkerStatus
            {
                State = resultingState,
                IsModuleAvailable = success,
                IsConsoleVisible = false,
                LastError = lastError
            };
        }

        public override Task<bool> StartWorkerAsync(CancellationToken cancellationToken = default)
        {
            return Task.FromResult(_startWorkerResult);
        }

        public override Task<Result<SetWorkerConsoleVisibilityResponse>> SetWorkerConsoleVisibilityAsync(
            bool isVisible,
            Action<EventEnvelope>? eventHandler = null,
            CancellationToken cancellationToken = default)
        {
            LastRequestedVisibility = isVisible;

            if (_toggleTask != null)
            {
                return _toggleTask;
            }

            if (_toggleResult.IsSuccess && _toggleResult.Value != null)
            {
                _status = new WorkerStatus
                {
                    State = WorkerConnectionState.Connected,
                    IsModuleAvailable = true,
                    IsConsoleVisible = _toggleResult.Value.IsVisible
                };
            }

            return Task.FromResult(_toggleResult);
        }
    }

    private sealed class ConnectWorkerServiceStub : WorkerConsoleConnectionWorkerServiceStub
    {
        public bool ConnectInvoked { get; private set; }

        public override Task<Result<ConnectionStatusDto>> ConnectExchangeAsync(Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            ConnectInvoked = true;

            return Task.FromResult(Result<ConnectionStatusDto>.Success(new ConnectionStatusDto
            {
                State = OnlyExo365.Contracts.Dtos.ConnectionState.Connected,
                UserPrincipalName = "admin@contoso.com",
                Organization = "contoso.onmicrosoft.com",
                GraphConnected = true
            }));
        }
    }

    private sealed class InteractiveBootstrapServiceStub : IInteractiveExchangeBootstrapService
    {
        public bool WasInvoked { get; private set; }
        public Result Result { get; set; } = Result.Success();

        public Task<Result> EnsureReadyAsync(Action<LogLevel, string>? onLog = null, CancellationToken cancellationToken = default)
        {
            WasInvoked = true;
            return Task.FromResult(Result);
        }
    }
}

