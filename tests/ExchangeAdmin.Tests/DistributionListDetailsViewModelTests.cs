using System.Reflection;
using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;
using ExchangeAdmin.Infrastructure.Ipc;
using ExchangeAdmin.Presentation.Services;
using ExchangeAdmin.Presentation.ViewModels;

namespace ExchangeAdmin.Tests;

public sealed class DistributionListDetailsViewModelTests
{
    [Fact]
    public async Task LoadDetailsAsync_PopulatesDynamicMembersReturnedByWorker()
    {
        var worker = new DistributionListDetailsWorkerService();
        var navigationService = new NavigationService();
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), navigationService);
        SetExchangeConnected(shell);
        SetCapabilities(shell, new CapabilityMapDto
        {
            Features = new FeatureCapabilitiesDto
            {
                CanGetDynamicDistributionGroup = true,
                CanGetDynamicDistributionGroupMember = true,
                CanGetRecipient = true
            }
        });

        var selectedItem = new DistributionListItemDto
        {
            Identity = "dynamic@contoso.com",
            DisplayName = "Dynamic Group",
            GroupType = "Dynamic"
        };

        var viewModel = new DistributionListDetailsViewModel(
            worker,
            navigationService,
            shell,
            () => selectedItem,
            _ => { },
            _ => { });

        await viewModel.LoadDetailsAsync(selectedItem.Identity, CancellationToken.None);

        Assert.True(viewModel.IsDynamicGroup);
        Assert.Equal(2, viewModel.Members.Count);
        Assert.True(viewModel.MembersHasMore);
        Assert.Equal("2 of 3 members", viewModel.MembersStatusText);
        Assert.Equal("dynamic@contoso.com", worker.LastDetailsRequest?.Identity);
        Assert.NotNull(worker.LastDetailsRequest);
        Assert.True(worker.LastDetailsRequest!.IncludeMembers);
    }

    [Fact]
    public async Task PreviewDynamicMembersCommand_PreservesPagingStateForLoadMore()
    {
        var worker = new DistributionListDetailsWorkerService
        {
            PreviewResponse = new PreviewDynamicGroupMembersResponse
            {
                Identity = "dynamic@contoso.com",
                TotalCount = 5,
                Members =
                [
                    new GroupMemberDto { Identity = "preview-1", Name = "Preview 1", PrimarySmtpAddress = "preview1@contoso.com" },
                    new GroupMemberDto { Identity = "preview-2", Name = "Preview 2", PrimarySmtpAddress = "preview2@contoso.com" }
                ]
            }
        };

        var navigationService = new NavigationService();
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), navigationService);
        SetExchangeConnected(shell);
        SetCapabilities(shell, new CapabilityMapDto
        {
            Features = new FeatureCapabilitiesDto
            {
                CanGetDynamicDistributionGroup = true,
                CanGetDynamicDistributionGroupMember = true
            }
        });

        var selectedItem = new DistributionListItemDto
        {
            Identity = "dynamic@contoso.com",
            DisplayName = "Dynamic Group",
            GroupType = "Dynamic"
        };

        var viewModel = new DistributionListDetailsViewModel(
            worker,
            navigationService,
            shell,
            () => selectedItem,
            _ => { },
            _ => { });

        await viewModel.LoadDetailsAsync(selectedItem.Identity, CancellationToken.None);
        Assert.True(viewModel.PreviewDynamicMembersCommand.CanExecute(null));

        viewModel.PreviewDynamicMembersCommand.Execute(null);
        await WaitForConditionAsync(() => worker.PreviewCalls == 1 && viewModel.Members.Count == 2 && viewModel.Members[0].Name == "Preview 1");

        Assert.True(viewModel.MembersHasMore);
        Assert.Equal("2 of 5 members", viewModel.MembersStatusText);
    }

    private static async Task WaitForConditionAsync(Func<bool> condition)
    {
        for (var attempt = 0; attempt < 80; attempt++)
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

    private static void SetCapabilities(ShellViewModel shell, CapabilityMapDto capabilities)
    {
        var property = typeof(ShellViewModel)
            .GetProperty(nameof(ShellViewModel.Capabilities), BindingFlags.Instance | BindingFlags.Public);

        Assert.NotNull(property);
        property!.SetValue(shell, capabilities);
    }

    private sealed class DistributionListDetailsWorkerService : TestDistributionListsWorkerServiceBase
    {
        public GetDistributionListDetailsRequest? LastDetailsRequest { get; private set; }
        public int PreviewCalls { get; private set; }
        public PreviewDynamicGroupMembersResponse PreviewResponse { get; set; } = new()
        {
            Identity = "dynamic@contoso.com",
            TotalCount = 3,
            Members =
            [
                new GroupMemberDto { Identity = "member-1", Name = "Member 1", PrimarySmtpAddress = "member1@contoso.com" },
                new GroupMemberDto { Identity = "member-2", Name = "Member 2", PrimarySmtpAddress = "member2@contoso.com" }
            ]
        };

        public override Task<Result<DistributionListDetailsDto>> GetDistributionListDetailsAsync(GetDistributionListDetailsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            LastDetailsRequest = request;
            return Task.FromResult(Result<DistributionListDetailsDto>.Success(new DistributionListDetailsDto
            {
                Identity = request.Identity,
                DisplayName = "Dynamic Group",
                PrimarySmtpAddress = "dynamic@contoso.com",
                GroupType = "Dynamic",
                IsDynamic = true,
                Members = new GroupMembersPageDto
                {
                    TotalCount = 3,
                    PageSize = 2,
                    HasMore = true,
                    Members =
                    [
                        new GroupMemberDto { Identity = "member-1", Name = "Member 1", PrimarySmtpAddress = "member1@contoso.com" },
                        new GroupMemberDto { Identity = "member-2", Name = "Member 2", PrimarySmtpAddress = "member2@contoso.com" }
                    ]
                }
            }));
        }

        public override Task<Result<PreviewDynamicGroupMembersResponse>> PreviewDynamicGroupMembersAsync(PreviewDynamicGroupMembersRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
        {
            PreviewCalls++;
            return Task.FromResult(Result<PreviewDynamicGroupMembersResponse>.Success(PreviewResponse));
        }
    }
}
