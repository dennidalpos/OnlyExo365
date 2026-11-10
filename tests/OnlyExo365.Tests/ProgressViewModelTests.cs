using System.Reflection;
using System.ComponentModel;
using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;
using OnlyExo365.Shell.Ipc;
using OnlyExo365.Shell.ViewModels;

namespace OnlyExo365.Tests;

public class ProgressViewModelTests
{
    [Fact]
    public void ShellViewModel_ExposesGlobalItemProgressFromProgressEvents()
    {
        var worker = new ProgressTestWorkerService();
        var navigation = new NavigationService();
        using var shell = new ShellViewModel(worker, navigation);

        worker.EmitProgress("corr-1", 42, "Working...", 21, 50);

        Assert.Equal(42, shell.GlobalProgress);
        Assert.Equal("Working...", shell.GlobalStatus);
        Assert.True(shell.HasGlobalItemProgress);
        Assert.Equal(21, shell.GlobalCurrentItem);
        Assert.Equal(50, shell.GlobalTotalItems);
        Assert.Equal("Loaded: 21 | Remaining: 29", shell.GlobalProgressCountText);
        Assert.True(shell.IsGlobalOperationRunning);

        worker.EmitProgress("corr-1", 100, "Done", null, null);

        Assert.False(shell.IsGlobalOperationRunning);
        Assert.False(shell.HasGlobalItemProgress);
        Assert.Null(shell.GlobalProgressCountText);
    }

    [Fact]
    public void ShellViewModel_IgnoresBackgroundProgressEventsRegisteredByCorrelationId()
    {
        var worker = new ProgressTestWorkerService();
        using var shell = new ShellViewModel(worker, new NavigationService());

        shell.RegisterBackgroundProgressOperation("audit-queue-001");
        worker.EmitProgress("audit-queue-001", 35, "Background audit search...", 7, 20);

        Assert.False(shell.IsGlobalOperationRunning);
        Assert.Equal(0, shell.GlobalProgress);
        Assert.Null(shell.GlobalStatus);
    }

    [Fact]
    public void ShellViewModel_ShowsGraphDisconnectedWhenInitialBootstrapDidNotConnectGraph()
    {
        using var shell = new ShellViewModel(new ProgressTestWorkerService(), new NavigationService());
        SetExchangeConnected(shell);

        Assert.False(shell.IsGraphConnected);
        Assert.Equal("Disconnected", shell.GraphStateDisplay);
        Assert.Equal("#DCDCAA", shell.GraphStateColor);
        Assert.Contains("initial connection", shell.GraphStateTooltip, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ShellViewModel_LocksNavigationWhenRegisteredSourceReportsBlockingState()
    {
        using var shell = new ShellViewModel(new ProgressTestWorkerService(), new NavigationService());
        var source = new BlockingStateSource();

        shell.RegisterNavigationStateSource(
            source,
            () => source.IsBlocking,
            nameof(BlockingStateSource.IsBlocking));

        Assert.False(shell.IsNavigationLocked);
        Assert.True(shell.CanNavigate);

        source.IsBlocking = true;

        Assert.True(shell.IsNavigationLocked);
        Assert.False(shell.CanNavigate);

        source.IsBlocking = false;

        Assert.False(shell.IsNavigationLocked);
        Assert.True(shell.CanNavigate);
    }

    [Fact]
    public async Task DashboardViewModel_TracksNumericProgressFromWorkerEvents()
    {
        var worker = new ProgressTestWorkerService
        {
            DashboardStatsResult = Result<DashboardStatsDto>.Success(new DashboardStatsDto())
        };
        worker.DashboardPause = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var navigation = new NavigationService();
        using var shell = new ShellViewModel(worker, navigation);
        SetExchangeConnected(shell);
        var viewModel = new DashboardViewModel(worker, navigation, shell, new CacheService());

        var loadTask = viewModel.LoadAsync();

        Assert.Equal("Fetching licenses...", viewModel.LoadingStatus);
        Assert.True(viewModel.HasLoadingItemProgress);
        Assert.Equal(3, viewModel.LoadingCurrentItem);
        Assert.Equal(7, viewModel.LoadingTotalItems);
        Assert.Equal("Loaded: 3 | Remaining: 4", viewModel.LoadingCountText);
        Assert.NotNull(worker.LastDashboardRequest);
        Assert.True(worker.LastDashboardRequest!.IncludeUnifiedGroups);
        Assert.False(worker.LastDashboardRequest.QuickCount);

        worker.DashboardPause.SetResult();
        await loadTask;
    }

    [Fact]
    public async Task DashboardViewModel_RefreshesShellGraphStateAfterGraphBackedLoad()
    {
        var worker = new ProgressTestWorkerService
        {
            DashboardStatsResult = Result<DashboardStatsDto>.Success(new DashboardStatsDto()),
            ConnectionStatusResult = Result<ConnectionStatusDto>.Success(new ConnectionStatusDto
            {
                State = ConnectionState.Connected,
                UserPrincipalName = "admin@contoso.com",
                Organization = "contoso.onmicrosoft.com",
                GraphConnected = true,
                ComplianceConnected = true
            })
        };

        using var shell = new ShellViewModel(worker, new NavigationService());
        SetWorkerConnected(shell);
        SetExchangeConnected(shell);
        var viewModel = new DashboardViewModel(worker, new NavigationService(), shell, new CacheService());

        await viewModel.LoadAsync();

        Assert.True(shell.IsGraphConnected);
        Assert.Equal("Connected", shell.GraphStateDisplay);
    }

    [Fact]
    public async Task MailboxSpaceViewModel_FormatsMailboxCountsDuringScan()
    {
        var worker = new ProgressTestWorkerService
        {
            MailboxSpaceResult = Result<GetMailboxSpaceReportResponse>.Success(new GetMailboxSpaceReportResponse())
        };
        worker.MailboxSpacePause = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        using var shell = new ShellViewModel(worker, new NavigationService());
        SetExchangeConnected(shell);
        var viewModel = new MailboxSpaceViewModel(worker, new NavigationService(), shell);

        var scanTask = InvokeStartScanAsync(viewModel);

        Assert.Equal("Scanning...", viewModel.ProgressStatus);
        Assert.True(viewModel.HasProgressCount);
        Assert.Equal(4, viewModel.ProgressCurrentItem);
        Assert.Equal(10, viewModel.ProgressTotalItems);
        Assert.Equal("Loaded: 4 | Remaining: 6", viewModel.ProgressCountText);

        worker.MailboxSpacePause.SetResult();
        await scanTask;
    }

    [Fact]
    public async Task DashboardViewModel_ExposesWarningReferenceFromSuccessfulResult()
    {
        var worker = new ProgressTestWorkerService
        {
            DashboardStatsResult = Result<DashboardStatsDto>.Success(new DashboardStatsDto
            {
                CorrelationId = "dashboard-warning",
                Warnings = ["Mailbox counts are approximate due to fallback limits."],
                HasPartialData = true
            }, "dashboard-warning")
        };

        using var shell = new ShellViewModel(worker, new NavigationService());
        SetExchangeConnected(shell);
        var viewModel = new DashboardViewModel(worker, new NavigationService(), shell, new CacheService());

        await viewModel.LoadAsync();

        Assert.True(viewModel.HasWarnings);
        Assert.True(viewModel.HasDiagnosticReference);
        Assert.Equal("dashboard-warning", viewModel.DiagnosticCorrelationId);
        Assert.Contains("Ref:", viewModel.DiagnosticReferenceText);
    }

    [Fact]
    public async Task DashboardViewModel_ExposesAdminUserWarningsFromPartialGraphResult()
    {
        var worker = new ProgressTestWorkerService
        {
            DashboardStatsResult = Result<DashboardStatsDto>.Success(new DashboardStatsDto
            {
                CorrelationId = "dashboard-admin-partial",
                Warnings = ["Get-MgDirectoryRoleMember failed for role 'Exchange Administrator': throttled."],
                WarningDetails =
                [
                    new OperationWarningDto
                    {
                        Code = "AdminRoleMembersLoadFailed",
                        Scope = "Dashboard.AdminUsers",
                        Message = "Get-MgDirectoryRoleMember failed for role 'Exchange Administrator': throttled.",
                        IsPartialData = true,
                        SampleItems = ["Exchange Administrator"]
                    }
                ],
                HasPartialData = true
            }, "dashboard-admin-partial")
        };

        using var shell = new ShellViewModel(worker, new NavigationService());
        SetExchangeConnected(shell);
        var viewModel = new DashboardViewModel(worker, new NavigationService(), shell, new CacheService());

        await viewModel.LoadAsync();

        Assert.True(viewModel.HasWarnings);
        Assert.Equal("dashboard-admin-partial", viewModel.DiagnosticCorrelationId);
        Assert.Contains(viewModel.Warnings, warning => warning.Contains("Exchange Administrator", StringComparison.Ordinal));
    }

    [Fact]
    public async Task MailboxAccessReportViewModel_ExposesWarningsAndCorrelationId()
    {
        var worker = new ProgressTestWorkerService
        {
            MailboxAccessResult = Result<GetMailboxAccessReportResponse>.Success(new GetMailboxAccessReportResponse
            {
                CorrelationId = "mailbox-access-warning",
                Warnings = ["Get-MailboxPermission failed for 2 mailbox(es). The report is partial."],
                HasPartialData = true
            }, "mailbox-access-warning")
        };

        using var shell = new ShellViewModel(worker, new NavigationService());
        SetExchangeConnected(shell);
        var viewModel = new MailboxAccessReportViewModel(worker, shell);

        await viewModel.LoadAsync();

        Assert.True(viewModel.HasWarnings);
        Assert.Equal("mailbox-access-warning", viewModel.DiagnosticCorrelationId);
        Assert.Contains("partial", viewModel.WarningsText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task MessageTraceViewModel_ExposesWarningsAndCorrelationId()
    {
        var worker = new ProgressTestWorkerService
        {
            MessageTraceResult = Result<GetMessageTraceResponse>.Success(new GetMessageTraceResponse
            {
                CorrelationId = "message-trace-warning",
                Warnings = ["Message trace returned partial telemetry for the selected time range."],
                HasPartialData = true,
                Messages = new List<MessageTraceItemDto>()
            }, "message-trace-warning")
        };

        using var shell = new ShellViewModel(worker, new NavigationService());
        SetExchangeConnected(shell);
        var viewModel = new MessageTraceViewModel(worker, shell);

        await InvokeFetchMessagesAsync(viewModel);

        Assert.True(viewModel.HasWarnings);
        Assert.Equal("message-trace-warning", viewModel.DiagnosticCorrelationId);
        Assert.Contains("partial telemetry", viewModel.WarningsText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task MessageTraceViewModel_UsesExactTotalInResultSummary()
    {
        var worker = new ProgressTestWorkerService
        {
            MessageTraceResult = Result<GetMessageTraceResponse>.Success(new GetMessageTraceResponse
            {
                TotalCount = 321,
                IsTotalCountExact = true,
                Messages = new List<MessageTraceItemDto>
                {
                    new()
                    {
                        MessageId = "msg-1",
                        MessageTraceId = "trace-1",
                        SenderAddress = "sender@contoso.com",
                        RecipientAddress = "recipient@contoso.com",
                        Status = "Delivered"
                    }
                }
            })
        };

        using var shell = new ShellViewModel(worker, new NavigationService());
        SetExchangeConnected(shell);
        var viewModel = new MessageTraceViewModel(worker, shell);

        await InvokeFetchMessagesAsync(viewModel);

        Assert.True(viewModel.IsTotalCountExact);
        Assert.Equal(321, viewModel.TotalCount);
        Assert.Equal("Total messages: 321", viewModel.ResultSummaryText);
    }

    private static async Task InvokeStartScanAsync(MailboxSpaceViewModel viewModel)
    {
        var method = typeof(MailboxSpaceViewModel)
            .GetMethod("StartScanAsync", BindingFlags.Instance | BindingFlags.NonPublic);

        Assert.NotNull(method);
        var task = method!.Invoke(viewModel, new object[] { CancellationToken.None }) as Task;
        Assert.NotNull(task);
        await task!;
    }

    private static async Task InvokeFetchMessagesAsync(MessageTraceViewModel viewModel)
    {
        var method = typeof(MessageTraceViewModel)
            .GetMethod("FetchMessagesAsync", BindingFlags.Instance | BindingFlags.NonPublic);

        Assert.NotNull(method);
        var task = method!.Invoke(viewModel, new object[] { CancellationToken.None }) as Task;
        Assert.NotNull(task);
        await task!;
    }

    private static void SetExchangeConnected(ShellViewModel shell)
    {
        var exchangeStateProperty = typeof(ShellViewModel)
            .GetProperty(nameof(ShellViewModel.ExchangeState), BindingFlags.Instance | BindingFlags.Public);

        Assert.NotNull(exchangeStateProperty);
        var setter = exchangeStateProperty!.SetMethod;
        Assert.NotNull(setter);
        setter!.Invoke(shell, new object[] { ConnectionState.Connected });
    }

    private static void SetWorkerConnected(ShellViewModel shell)
    {
        var workerStateProperty = typeof(ShellViewModel)
            .GetProperty(nameof(ShellViewModel.WorkerState), BindingFlags.Instance | BindingFlags.Public);

        Assert.NotNull(workerStateProperty);
        var setter = workerStateProperty!.SetMethod;
        Assert.NotNull(setter);
        setter!.Invoke(shell, new object[] { OnlyExo365.Shell.Ipc.WorkerConnectionState.Connected });
    }

    private sealed class ProgressTestWorkerService : TestWorkerServiceBase
    {
        public Result<DashboardStatsDto> DashboardStatsResult { get; set; } = Result<DashboardStatsDto>.Success(new DashboardStatsDto());
        public Result<GetMailboxSpaceReportResponse> MailboxSpaceResult { get; set; } = Result<GetMailboxSpaceReportResponse>.Success(new GetMailboxSpaceReportResponse());
        public Result<GetMailboxAccessReportResponse> MailboxAccessResult { get; set; } = Result<GetMailboxAccessReportResponse>.Success(new GetMailboxAccessReportResponse());
        public Result<GetMessageTraceResponse> MessageTraceResult { get; set; } = Result<GetMessageTraceResponse>.Success(new GetMessageTraceResponse());
        public Result<ConnectionStatusDto> ConnectionStatusResult { get; set; } = Result<ConnectionStatusDto>.Success(new ConnectionStatusDto
        {
            State = OnlyExo365.Contracts.Dtos.ConnectionState.Connected
        });
        public TaskCompletionSource? DashboardPause { get; set; }
        public TaskCompletionSource? MailboxSpacePause { get; set; }
        public GetDashboardStatsRequest? LastDashboardRequest { get; private set; }

        public void EmitProgress(string correlationId, int percent, string status, int? currentItem, int? totalItems)
        {
            PublishEventReceived(new EventEnvelope
            {
                CorrelationId = correlationId,
                EventType = EventType.Progress,
                Payload = JsonMessageSerializer.ToJsonElement(new ProgressEventPayload
                {
                    PercentComplete = percent,
                    StatusMessage = status,
                    CurrentItem = currentItem,
                    TotalItems = totalItems
                })
            });
        }

        public override Task<Result<ConnectionStatusDto>> GetConnectionStatusAsync(CancellationToken cancellationToken = default)
            => Task.FromResult(ConnectionStatusResult);

        public override Task<Result<DashboardStatsDto>> GetDashboardStatsAsync(GetDashboardStatsRequest? request = null, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            LastDashboardRequest = request;
            eventHandler?.Invoke(
                new EventEnvelope
                {
                    CorrelationId = "dashboard-progress",
                    EventType = EventType.Progress,
                    Payload = JsonMessageSerializer.ToJsonElement(new ProgressEventPayload
                    {
                        PercentComplete = 80,
                        StatusMessage = "Fetching licenses...",
                        CurrentItem = 3,
                        TotalItems = 7
                    })
                });

            return WaitForResultAsync(DashboardPause, DashboardStatsResult);
        }

        public override Task<Result<GetMailboxSpaceReportResponse>> GetMailboxSpaceReportAsync(GetMailboxSpaceReportRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            eventHandler?.Invoke(
                new EventEnvelope
                {
                    CorrelationId = "mailbox-space-progress",
                    EventType = EventType.Progress,
                    Payload = JsonMessageSerializer.ToJsonElement(new ProgressEventPayload
                    {
                        PercentComplete = 40,
                        StatusMessage = "Scanning...",
                        CurrentItem = 4,
                        TotalItems = 10
                    })
                });

            return WaitForResultAsync(MailboxSpacePause, MailboxSpaceResult);
        }

        public override Task<Result<GetMailboxAccessReportResponse>> GetMailboxAccessReportAsync(GetMailboxAccessReportRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(MailboxAccessResult);
        public override Task<Result<GetMessageTraceResponse>> GetMessageTraceAsync(GetMessageTraceRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(MessageTraceResult);

        private static async Task<Result<T>> WaitForResultAsync<T>(TaskCompletionSource? gate, Result<T> result)
        {
            if (gate != null)
            {
                await gate.Task;
            }

            return result;
        }
    }

    private sealed class BlockingStateSource : INotifyPropertyChanged
    {
        private bool _isBlocking;

        public bool IsBlocking
        {
            get => _isBlocking;
            set
            {
                if (_isBlocking == value)
                {
                    return;
                }

                _isBlocking = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsBlocking)));
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
    }
}

