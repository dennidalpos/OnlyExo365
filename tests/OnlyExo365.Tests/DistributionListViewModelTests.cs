using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;
using OnlyExo365.Shell.Ipc;
using OnlyExo365.Shell.Text;
using OnlyExo365.Shell.Services;
using OnlyExo365.Shell.ViewModels;

namespace OnlyExo365.Tests;

public sealed class DistributionListViewModelTests
{
    [Fact]
    public async Task LoadAsync_WhenDisconnected_DoesNotRequestAcceptedDomains()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        var worker = new DistributionListWorkerService();
        var viewModel = new DistributionListViewModel(worker, new NavigationService(), shell);

        await viewModel.LoadAsync();

        Assert.Equal(0, worker.GetAcceptedDomainsCalls);
        Assert.Empty(viewModel.AvailableMailDomains);
        Assert.Null(viewModel.ErrorMessage);
        Assert.True(shell.GlobalAlert.IsVisible);
        Assert.Equal(UserMessageCatalog.ConnectionRequiredAlertTitle, shell.GlobalAlert.Title);
        Assert.Equal(UserMessageCatalog.ConnectionRequiredAlertMessage, shell.GlobalAlert.Message);
    }

    private sealed class DistributionListWorkerService : TestDistributionListsWorkerServiceBase
    {
        public int GetAcceptedDomainsCalls { get; private set; }

        public override Task<Result<GetAcceptedDomainsResponse>> GetAcceptedDomainsAsync(
            GetAcceptedDomainsRequest request,
            Action<EventEnvelope>? eventHandler = null,
            CancellationToken cancellationToken = default)
        {
            GetAcceptedDomainsCalls++;
            return Task.FromResult(Result<GetAcceptedDomainsResponse>.Success(new GetAcceptedDomainsResponse
            {
                Domains = [new AcceptedDomainDto { DomainName = "contoso.com", Default = true }]
            }));
        }

        public override Task<Result<GetDistributionListsResponse>> GetDistributionListsAsync(
            GetDistributionListsRequest request,
            Action<EventEnvelope>? eventHandler = null,
            CancellationToken cancellationToken = default)
        {
            return Task.FromResult(Result<GetDistributionListsResponse>.Success(new GetDistributionListsResponse()));
        }
    }
}

