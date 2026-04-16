using System.Reflection;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;
using ExchangeAdmin.Infrastructure.Ipc;
using ExchangeAdmin.Presentation.Services;
using ExchangeAdmin.Presentation.ViewModels;

namespace ExchangeAdmin.Tests;

public sealed class ComplianceViewModelTests
{
    [Fact]
    public async Task LoadAsync_PopulatesWorkspaceCollections()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var worker = new ComplianceTestWorkerService();
        var viewModel = new ComplianceViewModel(worker, shell);

        await viewModel.LoadAsync();

        Assert.Single(viewModel.Searches);
        Assert.Single(viewModel.Cases);
        Assert.Single(viewModel.Actions);
        Assert.Equal(1, worker.GetComplianceWorkspaceCalls);
    }

    [Fact]
    public void SelectingSearch_PrefillsHoldName()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var viewModel = new ComplianceViewModel(new ComplianceTestWorkerService(), shell);

        viewModel.SelectedSearch = new ComplianceSearchDto
        {
            Name = "Incident Search"
        };

        Assert.Equal("Incident Search Hold", viewModel.HoldName);
    }

    [Fact]
    public async Task LoadAsync_ExposesWarningsAndDiagnosticReferenceFromPartialWorkspace()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var worker = new ComplianceTestWorkerService
        {
            WorkspaceResult = Result<GetComplianceWorkspaceResponse>.Success(new GetComplianceWorkspaceResponse
            {
                CorrelationId = "compliance-warning",
                Warnings = ["Get-CaseHoldPolicy failed for 2 case(s). The Compliance workspace remains partial."],
                HasPartialData = true,
                IsHoldListingUnsupported = true,
                HoldListingStatusMessage = "Existing holds are not visible in this Purview session.",
                Searches =
                [
                    new ComplianceSearchDto { Name = "Search-01" }
                ]
            }, "compliance-warning")
        };
        var viewModel = new ComplianceViewModel(worker, shell);

        await viewModel.LoadAsync();

        Assert.True(viewModel.HasWarnings);
        Assert.True(viewModel.HasDiagnosticReference);
        Assert.True(viewModel.HasHoldListingStatus);
        Assert.Equal("compliance-warning", viewModel.DiagnosticCorrelationId);
        Assert.Contains("Compliance workspace remains partial", viewModel.WarningsText, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Purview", viewModel.HoldListingStatusMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task SearchAuditLog_QueuesTasksWithoutBlockingNavigation()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var worker = new ComplianceTestWorkerService();
        var firstSearchGate = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var firstCorrelationId = "audit-001";
        var secondCorrelationId = "audit-002";
        var callIndex = 0;

        worker.SearchHandler = async (_, eventHandler, _) =>
        {
            callIndex++;
            if (callIndex == 1)
            {
                eventHandler?.Invoke(CreateProgressEvent(firstCorrelationId, 15, "Running first audit job..."));
                await firstSearchGate.Task;
                eventHandler?.Invoke(CreateProgressEvent(firstCorrelationId, 100, "First audit job complete"));
                return Result<SearchUnifiedAuditLogResponse>.Success(new SearchUnifiedAuditLogResponse
                {
                    Results =
                    [
                        new UnifiedAuditLogRecordDto { Identity = "first-record" }
                    ],
                    TotalCount = 1
                }, firstCorrelationId);
            }

            eventHandler?.Invoke(CreateProgressEvent(secondCorrelationId, 50, "Running second audit job..."));
            eventHandler?.Invoke(CreateProgressEvent(secondCorrelationId, 100, "Second audit job complete"));
            return Result<SearchUnifiedAuditLogResponse>.Success(new SearchUnifiedAuditLogResponse
            {
                Results =
                [
                    new UnifiedAuditLogRecordDto { Identity = "second-record" }
                ],
                TotalCount = 1
            }, secondCorrelationId);
        };

        var viewModel = new ComplianceViewModel(worker, shell);
        shell.RegisterNavigationStateSource(
            viewModel,
            () => viewModel.IsBusy,
            nameof(ComplianceViewModel.IsLoadingWorkspace),
            nameof(ComplianceViewModel.IsSearchingAudit),
            nameof(ComplianceViewModel.IsCreatingSearch),
            nameof(ComplianceViewModel.IsApplyingAction));

        viewModel.SearchAuditLogCommand.Execute(null);
        viewModel.SearchAuditLogCommand.Execute(null);

        await WaitForConditionAsync(() =>
            viewModel.AuditSearchTasks.Count == 2 &&
            viewModel.AuditSearchTasks[0].IsRunning &&
            viewModel.AuditSearchTasks[1].IsQueued);

        Assert.False(viewModel.IsBusy);
        Assert.True(viewModel.IsSearchingAudit);
        Assert.False(shell.IsNavigationLocked);
        Assert.True(viewModel.CanSearchAuditLog);

        firstSearchGate.SetResult();

        await WaitForConditionAsync(() => viewModel.AuditSearchTasks.All(task => task.IsCompleted));

        Assert.Equal(2, viewModel.AuditSearchTasks.Count);
        Assert.Equal("second-record", viewModel.AuditResults.Single().Identity);
        Assert.Equal(2, worker.SearchRequests.Count);
    }

    [Fact]
    public async Task SelectingAuditTask_UpdatesVisibleResults()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var worker = new ComplianceTestWorkerService();
        var callIndex = 0;

        worker.SearchHandler = (_, eventHandler, _) =>
        {
            callIndex++;
            var correlationId = $"audit-{callIndex:000}";
            eventHandler?.Invoke(CreateProgressEvent(correlationId, 100, $"Audit job {callIndex} complete"));

            return Task.FromResult(Result<SearchUnifiedAuditLogResponse>.Success(new SearchUnifiedAuditLogResponse
            {
                Results =
                [
                    new UnifiedAuditLogRecordDto { Identity = $"record-{callIndex}" }
                ],
                TotalCount = 1
            }, correlationId));
        };

        var viewModel = new ComplianceViewModel(worker, shell);

        viewModel.SearchAuditLogCommand.Execute(null);
        viewModel.SearchAuditLogCommand.Execute(null);

        await WaitForConditionAsync(() => viewModel.AuditSearchTasks.Count == 2 && viewModel.AuditSearchTasks.All(task => task.IsCompleted));

        viewModel.SelectedAuditSearchTask = viewModel.AuditSearchTasks[0];
        Assert.Equal("record-1", viewModel.AuditResults.Single().Identity);

        viewModel.SelectedAuditSearchTask = viewModel.AuditSearchTasks[1];
        Assert.Equal("record-2", viewModel.AuditResults.Single().Identity);
    }

    private static EventEnvelope CreateProgressEvent(string correlationId, int percentComplete, string statusMessage)
    {
        return new EventEnvelope
        {
            CorrelationId = correlationId,
            EventType = EventType.Progress,
            Payload = JsonMessageSerializer.ToJsonElement(new ProgressEventPayload
            {
                PercentComplete = percentComplete,
                StatusMessage = statusMessage
            })
        };
    }

    private static void SetExchangeConnected(ShellViewModel shell)
    {
        var property = typeof(ShellViewModel)
            .GetProperty(nameof(ShellViewModel.ExchangeState), BindingFlags.Instance | BindingFlags.Public);

        Assert.NotNull(property);
        property!.SetValue(shell, ConnectionState.Connected);
    }

    private static async Task WaitForConditionAsync(Func<bool> predicate)
    {
        for (var attempt = 0; attempt < 100; attempt++)
        {
            if (predicate())
            {
                return;
            }

            await Task.Delay(20);
        }

        Assert.True(predicate(), "Condition was not reached within the timeout.");
    }

    private sealed class ComplianceTestWorkerService : TestComplianceWorkerServiceBase
    {
        public int GetComplianceWorkspaceCalls { get; private set; }
        public List<SearchUnifiedAuditLogRequest> SearchRequests { get; } = new();

        public Result<GetComplianceWorkspaceResponse> WorkspaceResult { get; set; } = Result<GetComplianceWorkspaceResponse>.Success(new GetComplianceWorkspaceResponse
        {
            Searches =
            [
                new ComplianceSearchDto
                {
                    Name = "Search-01",
                    CaseName = "Case-01",
                    Status = "Completed",
                    ExchangeLocations = ["user@contoso.com"],
                    ContentMatchQuery = "kind:email"
                }
            ],
            Cases =
            [
                new ComplianceCaseDto
                {
                    Name = "Case-01",
                    Status = "Active",
                    CaseType = "eDiscovery"
                }
            ],
            Actions =
            [
                new ComplianceActionSummaryDto
                {
                    Name = "Search-01_Purge",
                    ActionType = "Purge",
                    SearchName = "Search-01",
                    Status = "Submitted"
                }
            ]
        });

        public Func<SearchUnifiedAuditLogRequest, Action<EventEnvelope>?, CancellationToken, Task<Result<SearchUnifiedAuditLogResponse>>>? SearchHandler { get; set; }

        public override Task<Result<GetComplianceWorkspaceResponse>> GetComplianceWorkspaceAsync(GetComplianceWorkspaceRequest? request = null, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            GetComplianceWorkspaceCalls++;
            return Task.FromResult(WorkspaceResult);
        }

        public override Task<Result<SearchUnifiedAuditLogResponse>> SearchUnifiedAuditLogAsync(SearchUnifiedAuditLogRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            SearchRequests.Add(request);

            if (SearchHandler != null)
            {
                return SearchHandler(request, eventHandler, cancellationToken);
            }

            return Task.FromResult(Result<SearchUnifiedAuditLogResponse>.Success(new SearchUnifiedAuditLogResponse
            {
                Results =
                [
                    new UnifiedAuditLogRecordDto { Identity = "default-record" }
                ],
                TotalCount = 1
            }, "audit-default"));
        }

        public override Task<Result> CreateComplianceSearchAsync(CreateComplianceSearchRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => throw TestStubExceptions.CreateUnsupported();

        public override Task<Result> StartComplianceSearchAsync(StartComplianceSearchRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => throw TestStubExceptions.CreateUnsupported();

        public override Task<Result> RemoveComplianceSearchAsync(RemoveComplianceSearchRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => throw TestStubExceptions.CreateUnsupported();

        public override Task<Result<InvokeComplianceActionResponse>> InvokeComplianceActionAsync(InvokeComplianceActionRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => throw TestStubExceptions.CreateUnsupported();
    }
}
