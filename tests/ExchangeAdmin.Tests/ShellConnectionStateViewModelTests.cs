using ExchangeAdmin.Contracts.Diagnostics;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Errors;
using ExchangeAdmin.Domain.Results;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts;
using ExchangeAdmin.Infrastructure.Ipc;
using ExchangeAdmin.Presentation.ViewModels;

namespace ExchangeAdmin.Tests;

public sealed class ShellConnectionStateViewModelTests : IDisposable
{
    private readonly string _tempDirectory = Path.Combine(Path.GetTempPath(), "OnlyExo365.Tests", Guid.NewGuid().ToString("N"));

    [Fact]
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

    [Fact]
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

    [Fact]
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

    [Fact]
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

    [Fact]
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
        }

        public override Task<Result<SetWorkerConsoleVisibilityResponse>> SetWorkerConsoleVisibilityAsync(
            bool isVisible,
            Action<EventEnvelope>? eventHandler = null,
            CancellationToken cancellationToken = default)
        {
            LastRequestedVisibility = isVisible;

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
                State = ExchangeAdmin.Contracts.Dtos.ConnectionState.Connected,
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
