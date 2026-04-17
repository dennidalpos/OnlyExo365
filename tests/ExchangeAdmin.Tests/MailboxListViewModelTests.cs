using System.Reflection;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;
using ExchangeAdmin.Infrastructure.Ipc;
using ExchangeAdmin.Presentation.Services;
using ExchangeAdmin.Presentation.ViewModels;

namespace ExchangeAdmin.Tests;

public sealed class MailboxListViewModelTests
{
    [Fact]
    public async Task CreateMailboxCommand_ClearsPasswordAfterDispatch()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var worker = new MailboxListTestWorkerService();
        var viewModel = new MailboxListViewModel(worker, new NavigationService(), shell)
        {
            NewMailboxDisplayName = "Helpdesk",
            NewMailboxAlias = "helpdesk",
            NewMailboxLocalPart = "helpdesk",
            SelectedMailboxDomain = "contoso.com"
        };

        viewModel.SetNewMailboxPassword("Sup3rSecret!");

        viewModel.CreateMailboxCommand.Execute(null);
        await WaitForConditionAsync(() => worker.CreateMailboxCalls == 1);

        Assert.Equal("Sup3rSecret!", worker.LastCreateMailboxRequest?.Password);
        Assert.False(viewModel.HasNewMailboxPassword);
        Assert.True(viewModel.NewMailboxPasswordClearTrigger > 0);
    }

    [Fact]
    public void SetRecipientTypeFilter_SharedMailbox_ClearsPendingPassword()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var viewModel = new MailboxListViewModel(new MailboxListTestWorkerService(), new NavigationService(), shell);

        viewModel.SetNewMailboxPassword("Sup3rSecret!");

        viewModel.SetRecipientTypeFilter("SharedMailbox", refresh: false);

        Assert.False(viewModel.HasNewMailboxPassword);
        Assert.True(viewModel.NewMailboxPasswordClearTrigger > 0);
    }

    [Fact]
    public async Task ShowProvisioningWorkspaceCommand_LoadsMemberUsersFromGraphView()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var worker = new MailboxListTestWorkerService();
        var viewModel = new MailboxListViewModel(worker, new NavigationService(), shell);

        viewModel.ShowProvisioningWorkspaceCommand.Execute(null);
        await WaitForConditionAsync(() => worker.GetMailboxProvisioningCandidatesCalls == 1);

        Assert.True(viewModel.IsProvisioningWorkspace);
        Assert.Single(viewModel.Provisioning.Candidates);
        Assert.True(worker.LastProvisioningRequest?.OnlyWithoutLicense);
        Assert.True(worker.LastProvisioningRequest?.OnlyWithoutMail);
        Assert.Equal(250, worker.LastProvisioningRequest?.PageSize);
    }

    [Fact]
    public async Task LoadAsync_UsesPageSize250ForMailboxPagination()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var worker = new MailboxListTestWorkerService();
        var viewModel = new MailboxListViewModel(worker, new NavigationService(), shell);

        await viewModel.LoadAsync();

        Assert.NotNull(worker.LastMailboxesRequest);
        Assert.Equal(250, worker.LastMailboxesRequest!.PageSize);
    }

    [Fact]
    public async Task LoadAsync_WhenDisconnected_DoesNotRequestAcceptedDomains()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        var worker = new MailboxListTestWorkerService();
        var viewModel = new MailboxListViewModel(worker, new NavigationService(), shell);

        await viewModel.LoadAsync();

        Assert.Equal(0, worker.GetAcceptedDomainsCalls);
        Assert.Empty(viewModel.AvailableMailDomains);
        Assert.Null(viewModel.ErrorMessage);
        Assert.True(shell.GlobalAlert.IsVisible);
    }

    [Fact]
    public async Task RefreshAsync_PreservesLoadedMailboxDepth()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var worker = new MailboxListTestWorkerService();
        worker.MailboxResponses.Enqueue(CreateMailboxResponse(250, totalCount: 600, hasMore: true));
        worker.MailboxResponses.Enqueue(CreateMailboxResponse(250, skip: 250, totalCount: 600, hasMore: true));
        worker.MailboxResponses.Enqueue(CreateMailboxResponse(500, totalCount: 600, hasMore: true));
        var viewModel = new MailboxListViewModel(worker, new NavigationService(), shell);

        await viewModel.LoadAsync();
        viewModel.LoadMoreCommand.Execute(null);
        await WaitForConditionAsync(() => worker.MailboxRequests.Count == 2);
        viewModel.RefreshCommand.Execute(null);
        await WaitForConditionAsync(() => worker.MailboxRequests.Count == 3);

        Assert.Equal([250, 250, 500], worker.MailboxRequests.Select(request => request.PageSize));
        Assert.Equal(500, viewModel.Mailboxes.Count);
        Assert.Equal(500, worker.LastMailboxesRequest?.PageSize);
    }

    private static async Task WaitForConditionAsync(Func<bool> condition)
    {
        for (var attempt = 0; attempt < 20; attempt++)
        {
            if (condition())
            {
                return;
            }

            await Task.Delay(25);
        }

        Assert.True(condition(), "Condition not reached in time.");
    }

    private static void SetExchangeConnected(ShellViewModel shell)
    {
        var property = typeof(ShellViewModel)
            .GetProperty(nameof(ShellViewModel.ExchangeState), BindingFlags.Instance | BindingFlags.Public);

        Assert.NotNull(property);
        property!.SetValue(shell, ConnectionState.Connected);
    }

    private sealed class MailboxListTestWorkerService : TestMailboxesWorkerServiceBase
    {
        public int CreateMailboxCalls { get; private set; }
        public CreateMailboxRequest? LastCreateMailboxRequest { get; private set; }
        public int GetMailboxProvisioningCandidatesCalls { get; private set; }
        public GetMailboxProvisioningCandidatesRequest? LastProvisioningRequest { get; private set; }
        public GetMailboxesRequest? LastMailboxesRequest { get; private set; }
        public int GetAcceptedDomainsCalls { get; private set; }
        public List<GetMailboxesRequest> MailboxRequests { get; } = [];
        public Queue<GetMailboxesResponse> MailboxResponses { get; } = new();

        public override Task<Result<GetMailboxProvisioningCandidatesResponse>> GetMailboxProvisioningCandidatesAsync(GetMailboxProvisioningCandidatesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            GetMailboxProvisioningCandidatesCalls++;
            LastProvisioningRequest = request;
            return Task.FromResult(Result<GetMailboxProvisioningCandidatesResponse>.Success(new GetMailboxProvisioningCandidatesResponse
            {
                Candidates =
                [
                    new MailboxProvisioningCandidateDto
                    {
                        DisplayName = "Mario Rossi",
                        UserPrincipalName = "mario.rossi@contoso.com",
                        Mail = null,
                        AccountEnabled = false,
                        HasAssignedLicense = false,
                        HasMailAddress = false
                    }
                ],
                TotalCount = 1,
                PageSize = request.PageSize,
                Skip = request.Skip,
                HasMore = false
            }));
        }

        public override Task<Result<GetMailboxesResponse>> GetMailboxesAsync(GetMailboxesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            LastMailboxesRequest = new GetMailboxesRequest
            {
                RecipientTypeDetails = request.RecipientTypeDetails,
                SearchQuery = request.SearchQuery,
                Filter = request.Filter,
                PageSize = request.PageSize,
                Skip = request.Skip,
                SortBy = request.SortBy,
                SortDescending = request.SortDescending
            };
            MailboxRequests.Add(LastMailboxesRequest);

            var response = MailboxResponses.Count > 0
                ? MailboxResponses.Dequeue()
                : CreateMailboxResponse(request.PageSize, request.Skip, totalCount: 1, hasMore: false);

            response.PageSize = request.PageSize;
            response.Skip = request.Skip;
            return Task.FromResult(Result<GetMailboxesResponse>.Success(response));
        }

        public override Task<Result> CreateMailboxAsync(CreateMailboxRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            CreateMailboxCalls++;
            LastCreateMailboxRequest = request;
            return Task.FromResult(Result.Success());
        }

        public override Task<Result<GetAcceptedDomainsResponse>> GetAcceptedDomainsAsync(GetAcceptedDomainsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            GetAcceptedDomainsCalls++;
            return Task.FromResult(Result<GetAcceptedDomainsResponse>.Success(new GetAcceptedDomainsResponse
            {
                Domains = [new AcceptedDomainDto { DomainName = "contoso.com", Default = true }]
            }));
        }

    }

    private static GetMailboxesResponse CreateMailboxResponse(int pageSize, int skip = 0, int totalCount = 1, bool hasMore = false)
    {
        var mailboxes = Enumerable.Range(skip, pageSize)
            .Select(index => new MailboxListItemDto
            {
                Identity = $"mailbox-{index:D4}",
                DisplayName = $"Mailbox {index:D4}",
                PrimarySmtpAddress = $"mailbox{index:D4}@contoso.com",
                RecipientType = "UserMailbox",
                RecipientTypeDetails = "UserMailbox"
            })
            .ToList();

        return new GetMailboxesResponse
        {
            Mailboxes = mailboxes,
            TotalCount = totalCount,
            PageSize = pageSize,
            Skip = skip,
            HasMore = hasMore,
            IsTotalCountExact = true
        };
    }
}
