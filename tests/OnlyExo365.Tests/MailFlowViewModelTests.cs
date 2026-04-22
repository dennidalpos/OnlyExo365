using System.Reflection;
using System.Text.Json;
using OnlyExo365.Shell.Services;
using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;
using OnlyExo365.Shell.Ipc;
using OnlyExo365.Shell.ViewModels;
using OnlyExo365.Worker.PowerShell;

namespace OnlyExo365.Tests;

public sealed class MailFlowViewModelTests
{
    [Fact]
    public void MapAddressListsResponse_FlagsUnsupportedSectionFromStructuredWarning()
    {
        var response = ExoAddressListCommands.MapAddressListsResponse([], [CreateStructuredWarning(
            "AddressListCmdletUnavailable",
            "MailFlow.AddressLists",
                    "Address Lists are not supported in the current Exchange session: Get-AddressList is not available.")]);

        Assert.True(response.IsUnsupported);
        Assert.True(response.HasPartialData);
        Assert.Single(response.Warnings);
        Assert.Contains("Get-AddressList", response.Warnings[0], StringComparison.Ordinal);
    }

    [Fact]
    public void MapOfflineAddressBooksResponse_FlagsUnsupportedSectionFromStructuredWarning()
    {
        var response = ExoOfflineAddressBookCommands.MapOfflineAddressBooksResponse([], [CreateStructuredWarning(
            "OfflineAddressBookCmdletUnavailable",
            "MailFlow.OfflineAddressBooks",
                    "Offline Address Books are not supported in the current Exchange session: Get-OfflineAddressBook is not available.")]);

        Assert.True(response.IsUnsupported);
        Assert.True(response.HasPartialData);
        Assert.Single(response.Warnings);
        Assert.Contains("Get-OfflineAddressBook", response.Warnings[0], StringComparison.Ordinal);
    }

    [Fact]
    public async Task LoadAsync_DegradesUnsupportedDirectorySectionsWithoutGlobalError()
    {
        using var shell = new ShellViewModel(new ConnectedConnectionWorkerServiceStub(), new NavigationService());
        SetExchangeConnected(shell);
        var worker = new MailFlowTestWorkerService
        {
            AddressListsResult = Result<GetAddressListsResponse>.Success(new GetAddressListsResponse
            {
            Warnings = ["Address Lists are not supported in the current Exchange session: Get-AddressList is not available."],
                HasPartialData = true,
                IsUnsupported = true
            }, "corr-address-lists"),
            OfflineAddressBooksResult = Result<GetOfflineAddressBooksResponse>.Success(new GetOfflineAddressBooksResponse
            {
            Warnings = ["Offline Address Books are not supported in the current Exchange session: Get-OfflineAddressBook is not available."],
                HasPartialData = true,
                IsUnsupported = true
            }, "corr-oab")
        };
        var viewModel = new MailFlowViewModel(worker, shell);

        await viewModel.LoadAsync();

        Assert.False(viewModel.HasError);
        Assert.True(viewModel.HasAddressListSectionWarning);
        Assert.True(viewModel.HasOfflineAddressBookSectionWarning);
        Assert.False(viewModel.IsAddressListSectionSupported);
        Assert.False(viewModel.IsOfflineAddressBookSectionSupported);
        Assert.Contains("Get-AddressList", viewModel.AddressListSectionWarningMessage, StringComparison.Ordinal);
        Assert.Contains("Get-OfflineAddressBook", viewModel.OfflineAddressBookSectionWarningMessage, StringComparison.Ordinal);
        Assert.Single(viewModel.AddressBookPolicies);
        Assert.Single(viewModel.SharingPolicies);
    }

    private static string CreateStructuredWarning(string code, string scope, string message)
    {
        return "__EA_WARN__" + JsonSerializer.Serialize(new OperationWarningDto
        {
            Code = code,
            Scope = scope,
            Message = message,
            IsPartialData = true
        });
    }

    private static void SetExchangeConnected(ShellViewModel shell)
    {
        var property = typeof(ShellViewModel)
            .GetProperty(nameof(ShellViewModel.ExchangeState), BindingFlags.Instance | BindingFlags.Public);

        Assert.NotNull(property);
        property!.SetValue(shell, ConnectionState.Connected);
    }

    private sealed class MailFlowTestWorkerService : TestMailFlowWorkerServiceBase
    {
        public Result<GetAddressListsResponse> AddressListsResult { get; set; } = Result<GetAddressListsResponse>.Success(new GetAddressListsResponse());
        public Result<GetOfflineAddressBooksResponse> OfflineAddressBooksResult { get; set; } = Result<GetOfflineAddressBooksResponse>.Success(new GetOfflineAddressBooksResponse());

        public override Task<Result<GetTransportRulesResponse>> GetTransportRulesAsync(GetTransportRulesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result<GetTransportRulesResponse>.Success(new GetTransportRulesResponse()));

        public override Task<Result> SetTransportRuleStateAsync(SetTransportRuleStateRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result.Success());

        public override Task<Result> UpsertTransportRuleAsync(UpsertTransportRuleRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result.Success());

        public override Task<Result> RemoveTransportRuleAsync(RemoveTransportRuleRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result.Success());

        public override Task<Result<TestTransportRuleResponse>> TestTransportRuleAsync(TestTransportRuleRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result<TestTransportRuleResponse>.Success(new TestTransportRuleResponse()));

        public override Task<Result<GetConnectorsResponse>> GetConnectorsAsync(GetConnectorsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result<GetConnectorsResponse>.Success(new GetConnectorsResponse()));

        public override Task<Result<GetAcceptedDomainsResponse>> GetAcceptedDomainsAsync(GetAcceptedDomainsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result<GetAcceptedDomainsResponse>.Success(new GetAcceptedDomainsResponse()));

        public override Task<Result> UpsertConnectorAsync(UpsertConnectorRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result.Success());

        public override Task<Result> RemoveConnectorAsync(RemoveConnectorRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result.Success());

        public override Task<Result> UpsertAcceptedDomainAsync(UpsertAcceptedDomainRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result.Success());

        public override Task<Result> RemoveAcceptedDomainAsync(RemoveAcceptedDomainRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result.Success());

        public override Task<Result<GetRemoteDomainsResponse>> GetRemoteDomainsAsync(GetRemoteDomainsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result<GetRemoteDomainsResponse>.Success(new GetRemoteDomainsResponse()));

        public override Task<Result> UpsertRemoteDomainAsync(UpsertRemoteDomainRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result.Success());

        public override Task<Result> RemoveRemoteDomainAsync(RemoveRemoteDomainRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result.Success());

        public override Task<Result<GetOrganizationRelationshipsResponse>> GetOrganizationRelationshipsAsync(GetOrganizationRelationshipsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result<GetOrganizationRelationshipsResponse>.Success(new GetOrganizationRelationshipsResponse()));

        public override Task<Result<GetAddressListsResponse>> GetAddressListsAsync(GetAddressListsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(AddressListsResult);

        public override Task<Result> UpsertAddressListAsync(UpsertAddressListRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result.Success());

        public override Task<Result> RemoveAddressListAsync(RemoveAddressListRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result.Success());

        public override Task<Result<GetAddressBookPoliciesResponse>> GetAddressBookPoliciesAsync(GetAddressBookPoliciesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result<GetAddressBookPoliciesResponse>.Success(new GetAddressBookPoliciesResponse
            {
                Policies =
                [
                    new AddressBookPolicyDto
                    {
                        Identity = "abp-01",
                        Name = "ABP 01",
                        GlobalAddressList = "\\Default Global Address List",
                        OfflineAddressBook = "\\Default Offline Address Book"
                    }
                ]
            }));

        public override Task<Result> UpsertAddressBookPolicyAsync(UpsertAddressBookPolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result.Success());

        public override Task<Result> RemoveAddressBookPolicyAsync(RemoveAddressBookPolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result.Success());

        public override Task<Result<GetOfflineAddressBooksResponse>> GetOfflineAddressBooksAsync(GetOfflineAddressBooksRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(OfflineAddressBooksResult);

        public override Task<Result> UpsertOfflineAddressBookAsync(UpsertOfflineAddressBookRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result.Success());

        public override Task<Result> RemoveOfflineAddressBookAsync(RemoveOfflineAddressBookRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result.Success());

        public override Task<Result<GetSharingPoliciesResponse>> GetSharingPoliciesAsync(GetSharingPoliciesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result<GetSharingPoliciesResponse>.Success(new GetSharingPoliciesResponse
            {
                Policies =
                [
                    new SharingPolicyDto
                    {
                        Identity = "sharing-01",
                        Name = "Partner Sharing",
                        Domains = ["contoso.com: CalendarSharingFreeBusyDetail"],
                        Enabled = true
                    }
                ]
            }));

        public override Task<Result> UpsertSharingPolicyAsync(UpsertSharingPolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result.Success());

        public override Task<Result> RemoveSharingPolicyAsync(RemoveSharingPolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result.Success());

        public override Task<Result> UpsertOrganizationRelationshipAsync(UpsertOrganizationRelationshipRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result.Success());

        public override Task<Result> RemoveOrganizationRelationshipAsync(RemoveOrganizationRelationshipRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default)
            => Task.FromResult(Result.Success());
    }
}

