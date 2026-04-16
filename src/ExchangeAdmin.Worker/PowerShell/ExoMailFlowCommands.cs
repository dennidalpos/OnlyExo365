using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

public class ExoMailFlowCommands
{
    private readonly ExoMessageTraceCommands _messageTraceCommands;
    private readonly ExoTransportRuleCommands _transportRuleCommands;
    private readonly ExoConnectorCommands _connectorCommands;
    private readonly ExoAcceptedDomainCommands _acceptedDomainCommands;
    private readonly ExoRemoteDomainCommands _remoteDomainCommands;
    private readonly ExoOrganizationRelationshipCommands _organizationRelationshipCommands;
    private readonly ExoAddressListCommands _addressListCommands;
    private readonly ExoAddressBookPolicyCommands _addressBookPolicyCommands;
    private readonly ExoOfflineAddressBookCommands _offlineAddressBookCommands;
    private readonly ExoSharingPolicyCommands _sharingPolicyCommands;

    public ExoMailFlowCommands(PowerShellEngine engine)
    {
        _messageTraceCommands = new ExoMessageTraceCommands(engine);
        _transportRuleCommands = new ExoTransportRuleCommands(engine);
        _connectorCommands = new ExoConnectorCommands(engine);
        _acceptedDomainCommands = new ExoAcceptedDomainCommands(engine);
        _remoteDomainCommands = new ExoRemoteDomainCommands(engine);
        _organizationRelationshipCommands = new ExoOrganizationRelationshipCommands(engine);
        _addressListCommands = new ExoAddressListCommands(engine);
        _addressBookPolicyCommands = new ExoAddressBookPolicyCommands(engine);
        _offlineAddressBookCommands = new ExoOfflineAddressBookCommands(engine);
        _sharingPolicyCommands = new ExoSharingPolicyCommands(engine);
    }

    public Task<GetMessageTraceResponse> GetMessageTraceAsync(
        GetMessageTraceRequest request,
        CancellationToken cancellationToken = default)
        => _messageTraceCommands.GetMessageTraceAsync(request, cancellationToken);

    public Task<GetMessageTraceDetailsResponse> GetMessageTraceDetailsAsync(
        GetMessageTraceDetailsRequest request,
        CancellationToken cancellationToken = default)
        => _messageTraceCommands.GetMessageTraceDetailsAsync(request, cancellationToken);

    public Task<GetTransportRulesResponse> GetTransportRulesAsync(CancellationToken cancellationToken = default)
        => _transportRuleCommands.GetTransportRulesAsync(cancellationToken);

    public Task SetTransportRuleStateAsync(
        SetTransportRuleStateRequest request,
        CancellationToken cancellationToken = default)
        => _transportRuleCommands.SetTransportRuleStateAsync(request, cancellationToken);

    public Task UpsertTransportRuleAsync(
        UpsertTransportRuleRequest request,
        CancellationToken cancellationToken = default)
        => _transportRuleCommands.UpsertTransportRuleAsync(request, cancellationToken);

    public Task RemoveTransportRuleAsync(
        RemoveTransportRuleRequest request,
        CancellationToken cancellationToken = default)
        => _transportRuleCommands.RemoveTransportRuleAsync(request, cancellationToken);

    public Task<TestTransportRuleResponse> TestTransportRuleAsync(
        TestTransportRuleRequest request,
        CancellationToken cancellationToken = default)
        => _transportRuleCommands.TestTransportRuleAsync(request, cancellationToken);

    public Task<GetConnectorsResponse> GetConnectorsAsync(CancellationToken cancellationToken = default)
        => _connectorCommands.GetConnectorsAsync(cancellationToken);

    public Task UpsertConnectorAsync(
        UpsertConnectorRequest request,
        CancellationToken cancellationToken = default)
        => _connectorCommands.UpsertConnectorAsync(request, cancellationToken);

    public Task RemoveConnectorAsync(
        RemoveConnectorRequest request,
        CancellationToken cancellationToken = default)
        => _connectorCommands.RemoveConnectorAsync(request, cancellationToken);

    public Task<GetAcceptedDomainsResponse> GetAcceptedDomainsAsync(CancellationToken cancellationToken = default)
        => _acceptedDomainCommands.GetAcceptedDomainsAsync(cancellationToken);

    public Task UpsertAcceptedDomainAsync(
        UpsertAcceptedDomainRequest request,
        CancellationToken cancellationToken = default)
        => _acceptedDomainCommands.UpsertAcceptedDomainAsync(request, cancellationToken);

    public Task RemoveAcceptedDomainAsync(
        RemoveAcceptedDomainRequest request,
        CancellationToken cancellationToken = default)
        => _acceptedDomainCommands.RemoveAcceptedDomainAsync(request, cancellationToken);

    public Task<GetRemoteDomainsResponse> GetRemoteDomainsAsync(CancellationToken cancellationToken = default)
        => _remoteDomainCommands.GetRemoteDomainsAsync(cancellationToken);

    public Task UpsertRemoteDomainAsync(
        UpsertRemoteDomainRequest request,
        CancellationToken cancellationToken = default)
        => _remoteDomainCommands.UpsertRemoteDomainAsync(request, cancellationToken);

    public Task RemoveRemoteDomainAsync(
        RemoveRemoteDomainRequest request,
        CancellationToken cancellationToken = default)
        => _remoteDomainCommands.RemoveRemoteDomainAsync(request, cancellationToken);

    public Task<GetOrganizationRelationshipsResponse> GetOrganizationRelationshipsAsync(CancellationToken cancellationToken = default)
        => _organizationRelationshipCommands.GetOrganizationRelationshipsAsync(cancellationToken);

    public Task UpsertOrganizationRelationshipAsync(
        UpsertOrganizationRelationshipRequest request,
        CancellationToken cancellationToken = default)
        => _organizationRelationshipCommands.UpsertOrganizationRelationshipAsync(request, cancellationToken);

    public Task RemoveOrganizationRelationshipAsync(
        RemoveOrganizationRelationshipRequest request,
        CancellationToken cancellationToken = default)
        => _organizationRelationshipCommands.RemoveOrganizationRelationshipAsync(request, cancellationToken);

    public Task<GetAddressListsResponse> GetAddressListsAsync(CancellationToken cancellationToken = default)
        => _addressListCommands.GetAddressListsAsync(cancellationToken);

    public Task UpsertAddressListAsync(
        UpsertAddressListRequest request,
        CancellationToken cancellationToken = default)
        => _addressListCommands.UpsertAddressListAsync(request, cancellationToken);

    public Task RemoveAddressListAsync(
        RemoveAddressListRequest request,
        CancellationToken cancellationToken = default)
        => _addressListCommands.RemoveAddressListAsync(request, cancellationToken);

    public Task<GetAddressBookPoliciesResponse> GetAddressBookPoliciesAsync(CancellationToken cancellationToken = default)
        => _addressBookPolicyCommands.GetAddressBookPoliciesAsync(cancellationToken);

    public Task UpsertAddressBookPolicyAsync(
        UpsertAddressBookPolicyRequest request,
        CancellationToken cancellationToken = default)
        => _addressBookPolicyCommands.UpsertAddressBookPolicyAsync(request, cancellationToken);

    public Task RemoveAddressBookPolicyAsync(
        RemoveAddressBookPolicyRequest request,
        CancellationToken cancellationToken = default)
        => _addressBookPolicyCommands.RemoveAddressBookPolicyAsync(request, cancellationToken);

    public Task<GetOfflineAddressBooksResponse> GetOfflineAddressBooksAsync(CancellationToken cancellationToken = default)
        => _offlineAddressBookCommands.GetOfflineAddressBooksAsync(cancellationToken);

    public Task UpsertOfflineAddressBookAsync(
        UpsertOfflineAddressBookRequest request,
        CancellationToken cancellationToken = default)
        => _offlineAddressBookCommands.UpsertOfflineAddressBookAsync(request, cancellationToken);

    public Task RemoveOfflineAddressBookAsync(
        RemoveOfflineAddressBookRequest request,
        CancellationToken cancellationToken = default)
        => _offlineAddressBookCommands.RemoveOfflineAddressBookAsync(request, cancellationToken);

    public Task<GetSharingPoliciesResponse> GetSharingPoliciesAsync(CancellationToken cancellationToken = default)
        => _sharingPolicyCommands.GetSharingPoliciesAsync(cancellationToken);

    public Task UpsertSharingPolicyAsync(
        UpsertSharingPolicyRequest request,
        CancellationToken cancellationToken = default)
        => _sharingPolicyCommands.UpsertSharingPolicyAsync(request, cancellationToken);

    public Task RemoveSharingPolicyAsync(
        RemoveSharingPolicyRequest request,
        CancellationToken cancellationToken = default)
        => _sharingPolicyCommands.RemoveSharingPolicyAsync(request, cancellationToken);
}
