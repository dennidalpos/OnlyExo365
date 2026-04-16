using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;
using ExchangeAdmin.Infrastructure.Ipc;
using ExchangeAdmin.Presentation.Services;
using ExchangeAdmin.Presentation.ViewModels;

namespace ExchangeAdmin.Tests;

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
        Assert.Equal("Not connected to Exchange Online", viewModel.ErrorMessage);
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
