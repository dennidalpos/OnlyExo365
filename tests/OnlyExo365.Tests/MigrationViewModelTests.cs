using System.Reflection;
using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Errors;
using OnlyExo365.Contracts.Results;
using OnlyExo365.Shell.Ipc;
using OnlyExo365.Shell.ViewModels;

namespace OnlyExo365.Tests;

public sealed class MigrationViewModelTests
{
    [Fact]
    public async Task LoadAsync_PopulatesBatchesAndEndpoints()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var worker = new MigrationTestWorkerService();
        var viewModel = new MigrationViewModel(worker, shell);

        await viewModel.LoadAsync();

        Assert.Single(viewModel.Batches);
        Assert.Single(viewModel.Endpoints);
        Assert.Equal("batch-01", viewModel.SelectedBatch?.Identity);
        Assert.Equal("endpoint-01", viewModel.NewBatchEndpointIdentity);
        Assert.Null(viewModel.SelectedDetails);
        Assert.Equal(0, worker.GetMigrationBatchDetailsCalls);
        Assert.Equal(1, worker.GetMigrationBatchesCalls);
        Assert.Equal(1, worker.GetMigrationEndpointsCalls);
        Assert.Equal(250, worker.LastGetMigrationBatchesRequest?.PageSize);
    }

    [Fact]
    public async Task LoadBatchDetailsCommand_LoadsDetailsOnlyOnExplicitRequest()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var worker = new MigrationTestWorkerService();
        var viewModel = new MigrationViewModel(worker, shell);

        await viewModel.LoadAsync();

        viewModel.LoadBatchDetailsCommand.Execute(null);
        await worker.MigrationBatchDetailsRequested.Task.WaitAsync(TimeSpan.FromSeconds(2));

        Assert.Equal(1, worker.GetMigrationBatchDetailsCalls);
        Assert.Equal("batch-01", viewModel.SelectedDetails?.Identity);
        Assert.Contains("loaded on demand", viewModel.BatchDetailsStatusText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task SelectingEndpoint_PopulatesEditorAndBatchEndpoint()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var worker = new MigrationTestWorkerService();
        var viewModel = new MigrationViewModel(worker, shell);

        await viewModel.LoadAsync();

        viewModel.SelectedEndpoint = viewModel.Endpoints[0];

        Assert.Equal("endpoint-01", viewModel.NewBatchEndpointIdentity);
        Assert.Equal("Endpoint 01", viewModel.EndpointName);
        Assert.Equal("IMAP", viewModel.EndpointType);
        Assert.Equal("imap.contoso.com", viewModel.EndpointRemoteServer);
        Assert.Equal(993, viewModel.EndpointPort);
        Assert.True(viewModel.IsImapEndpointType);
    }

    [Fact]
    public async Task RunBatchPreflightAsync_UpdatesSummaryAndReadyState()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var worker = new MigrationTestWorkerService();
        var viewModel = new MigrationViewModel(worker, shell);
        await viewModel.LoadAsync();

        viewModel.NewBatchName = "IMAP Batch";
        viewModel.NewBatchType = "IMAP";
        viewModel.NewBatchEndpointIdentity = "endpoint-01";
        viewModel.NewBatchCsvFilePath = "C:\\temp\\migration.csv";

        viewModel.RunBatchPreflightCommand.Execute(null);
        await worker.MigrationBatchPreflightRequested.Task.WaitAsync(TimeSpan.FromSeconds(2));

        Assert.True(viewModel.IsBatchPreflightReady);
        Assert.Contains("Ready: True", viewModel.BatchPreflightSummary, StringComparison.Ordinal);
        Assert.Contains("Csv rows: 2", viewModel.BatchPreflightSummary, StringComparison.Ordinal);
    }

    [Fact]
    public async Task RunBatchPreflightAsync_UsesEnglishFallbackErrorWhenWorkerOmitsMessage()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var worker = new MigrationTestWorkerService { ReturnPreflightFailureWithoutMessage = true };
        var viewModel = new MigrationViewModel(worker, shell);
        await viewModel.LoadAsync();

        viewModel.NewBatchName = "IMAP Batch";
        viewModel.NewBatchType = "IMAP";
        viewModel.NewBatchEndpointIdentity = "endpoint-01";
        viewModel.NewBatchCsvFilePath = "C:\\temp\\migration.csv";

        viewModel.RunBatchPreflightCommand.Execute(null);
        await worker.MigrationBatchPreflightRequested.Task.WaitAsync(TimeSpan.FromSeconds(2));

        Assert.Equal("Unable to run migration preflight.", viewModel.ErrorMessage);
        Assert.DoesNotContain("eseguire", viewModel.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task TestEndpointCommand_ClearsPasswordAfterDispatch()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var worker = new MigrationTestWorkerService();
        var viewModel = new MigrationViewModel(worker, shell);

        viewModel.EndpointName = "Endpoint Draft";
        viewModel.EndpointType = "ExchangeRemoteMove";
        viewModel.EndpointRemoteServer = "mail.contoso.com";
        viewModel.EndpointUsername = "admin@contoso.com";
        viewModel.SetEndpointPassword("Sup3rSecret!");

        viewModel.TestEndpointCommand.Execute(null);
        await worker.MigrationEndpointTestRequested.Task.WaitAsync(TimeSpan.FromSeconds(2));

        Assert.Equal("Sup3rSecret!", worker.LastTestMigrationEndpointRequest?.Password);
        Assert.False(viewModel.HasEndpointPassword);
        Assert.True(viewModel.EndpointPasswordClearTrigger > 0);
    }

    [Fact]
    public async Task SelectingEndpoint_ClearsPendingPassword()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var worker = new MigrationTestWorkerService();
        var viewModel = new MigrationViewModel(worker, shell);
        await viewModel.LoadAsync();

        viewModel.SetEndpointPassword("Sup3rSecret!");

        viewModel.SelectedEndpoint = viewModel.Endpoints[0];

        Assert.False(viewModel.HasEndpointPassword);
        Assert.True(viewModel.EndpointPasswordClearTrigger > 0);
    }

    private static void SetExchangeConnected(ShellViewModel shell)
    {
        var property = typeof(ShellViewModel)
            .GetProperty(nameof(ShellViewModel.ExchangeState), BindingFlags.Instance | BindingFlags.Public);

        Assert.NotNull(property);
        property!.SetValue(shell, ConnectionState.Connected);
    }

    private sealed class MigrationTestWorkerService : TestMigrationWorkerServiceBase
    {
        public int GetMigrationBatchesCalls { get; private set; }
        public int GetMigrationEndpointsCalls { get; private set; }
        public int GetMigrationBatchDetailsCalls { get; private set; }
        public int GetMigrationBatchPreflightCalls { get; private set; }
        public int TestMigrationEndpointCalls { get; private set; }
        public bool ReturnPreflightFailureWithoutMessage { get; init; }
        public TestMigrationEndpointRequest? LastTestMigrationEndpointRequest { get; private set; }
        public GetMigrationBatchesRequest? LastGetMigrationBatchesRequest { get; private set; }
        public TaskCompletionSource MigrationBatchDetailsRequested { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public TaskCompletionSource MigrationBatchPreflightRequested { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public TaskCompletionSource MigrationEndpointTestRequested { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);

        public override Task<Result<GetMigrationBatchesResponse>> GetMigrationBatchesAsync(GetMigrationBatchesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            GetMigrationBatchesCalls++;
            LastGetMigrationBatchesRequest = new GetMigrationBatchesRequest
            {
                SearchQuery = request.SearchQuery,
                Status = request.Status,
                PageSize = request.PageSize,
                Skip = request.Skip,
                SortBy = request.SortBy,
                SortDescending = request.SortDescending
            };

            return Task.FromResult(Result<GetMigrationBatchesResponse>.Success(new GetMigrationBatchesResponse
            {
                Batches =
                [
                    new MigrationBatchListItemDto
                    {
                        Identity = "batch-01",
                        Name = "Batch 01",
                        Status = "Created",
                        BatchType = "IMAP",
                        SourceEndpoint = "endpoint-01"
                    }
                ],
                TotalCount = 1,
                PageSize = request.PageSize,
                Skip = request.Skip,
                HasMore = false
            }));
        }

        public override Task<Result<GetMigrationEndpointsResponse>> GetMigrationEndpointsAsync(GetMigrationEndpointsRequest? request = null, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            GetMigrationEndpointsCalls++;
            return Task.FromResult(Result<GetMigrationEndpointsResponse>.Success(new GetMigrationEndpointsResponse
            {
                Endpoints =
                [
                    new MigrationEndpointDto
                    {
                        Identity = "endpoint-01",
                        Name = "Endpoint 01",
                        EndpointType = "IMAP",
                        RemoteServer = "imap.contoso.com",
                        Port = 993,
                        Security = "Ssl"
                    }
                ]
            }));
        }

        public override Task<Result<MigrationBatchDetailsDto>> GetMigrationBatchDetailsAsync(GetMigrationBatchDetailsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            GetMigrationBatchDetailsCalls++;
            MigrationBatchDetailsRequested.TrySetResult();
            return Task.FromResult(Result<MigrationBatchDetailsDto>.Success(new MigrationBatchDetailsDto
            {
                Identity = request.Identity,
                Name = "Batch 01",
                Status = "Created",
                NotificationEmails = ["ops@contoso.com"]
            }));
        }

        public override Task<Result> UpsertMigrationEndpointAsync(UpsertMigrationEndpointRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result.Success());

        public override Task<Result<TestMigrationEndpointResponse>> TestMigrationEndpointAsync(TestMigrationEndpointRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            TestMigrationEndpointCalls++;
            LastTestMigrationEndpointRequest = request;
            MigrationEndpointTestRequested.TrySetResult();
            return Task.FromResult(Result<TestMigrationEndpointResponse>.Success(new TestMigrationEndpointResponse { Summary = "Endpoint test ok." }));
        }

        public override Task<Result<GetMigrationBatchPreflightResponse>> GetMigrationBatchPreflightAsync(GetMigrationBatchPreflightRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            GetMigrationBatchPreflightCalls++;
            MigrationBatchPreflightRequested.TrySetResult();
            if (ReturnPreflightFailureWithoutMessage)
            {
                return Task.FromResult(Result<GetMigrationBatchPreflightResponse>.Failure(
                    NormalizedError.Create(ErrorCode.Unknown, string.Empty)));
            }

            return Task.FromResult(Result<GetMigrationBatchPreflightResponse>.Success(new GetMigrationBatchPreflightResponse
            {
                IsReady = true,
                EndpointType = "IMAP",
                CsvRowCount = 2,
                CsvHeaders = ["EmailAddress", "UserName", "Password"],
                Messages = ["Preflight completed successfully."]
            }));
        }

        public override Task<Result> CreateMigrationBatchAsync(CreateMigrationBatchRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result.Success());

        public override Task<Result> StartMigrationBatchAsync(StartMigrationBatchRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result.Success());

        public override Task<Result> CompleteMigrationBatchAsync(CompleteMigrationBatchRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result.Success());

        public override Task<Result> RemoveMigrationBatchAsync(RemoveMigrationBatchRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result.Success());
    }
}

