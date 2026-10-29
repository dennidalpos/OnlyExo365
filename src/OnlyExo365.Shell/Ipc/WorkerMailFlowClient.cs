using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;

namespace OnlyExo365.Shell.Ipc;

internal sealed class WorkerMailFlowClient
{
    private readonly WorkerClientRuntime _runtime;

    public WorkerMailFlowClient(WorkerClientRuntime runtime)
    {
        _runtime = runtime;
    }

    public Task<Result<GetMessageTraceResponse>> GetMessageTraceAsync(
        GetMessageTraceRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetMessageTraceResponse>(OperationType.GetMessageTrace, request, eventHandler, cancellationToken);

    public Task<Result<GetMessageTraceDetailsResponse>> GetMessageTraceDetailsAsync(
        GetMessageTraceDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetMessageTraceDetailsResponse>(OperationType.GetMessageTraceDetails, request, eventHandler, cancellationToken);

    public Task<Result<GetTransportRulesResponse>> GetTransportRulesAsync(
        GetTransportRulesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetTransportRulesResponse>(OperationType.GetTransportRules, request, eventHandler, cancellationToken);

    public Task<Result> SetTransportRuleStateAsync(
        SetTransportRuleStateRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.SetTransportRuleState, request, eventHandler, cancellationToken);

    public Task<Result> UpsertTransportRuleAsync(
        UpsertTransportRuleRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.UpsertTransportRule, request, eventHandler, cancellationToken);

    public Task<Result> RemoveTransportRuleAsync(
        RemoveTransportRuleRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.RemoveTransportRule, request, eventHandler, cancellationToken);

    public Task<Result<TestTransportRuleResponse>> TestTransportRuleAsync(
        TestTransportRuleRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<TestTransportRuleResponse>(OperationType.TestTransportRule, request, eventHandler, cancellationToken);

    public Task<Result<GetConnectorsResponse>> GetConnectorsAsync(
        GetConnectorsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetConnectorsResponse>(OperationType.GetConnectors, request, eventHandler, cancellationToken);

    public Task<Result<GetAcceptedDomainsResponse>> GetAcceptedDomainsAsync(
        GetAcceptedDomainsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetAcceptedDomainsResponse>(OperationType.GetAcceptedDomains, request, eventHandler, cancellationToken);

    public Task<Result> UpsertConnectorAsync(
        UpsertConnectorRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.UpsertConnector, request, eventHandler, cancellationToken);

    public Task<Result> RemoveConnectorAsync(
        RemoveConnectorRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.RemoveConnector, request, eventHandler, cancellationToken);

    public Task<Result> UpsertAcceptedDomainAsync(
        UpsertAcceptedDomainRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.UpsertAcceptedDomain, request, eventHandler, cancellationToken);

    public Task<Result> RemoveAcceptedDomainAsync(
        RemoveAcceptedDomainRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.RemoveAcceptedDomain, request, eventHandler, cancellationToken);

    public Task<Result<GetRemoteDomainsResponse>> GetRemoteDomainsAsync(
        GetRemoteDomainsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetRemoteDomainsResponse>(OperationType.GetRemoteDomains, request, eventHandler, cancellationToken);

    public Task<Result> UpsertRemoteDomainAsync(
        UpsertRemoteDomainRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.UpsertRemoteDomain, request, eventHandler, cancellationToken);

    public Task<Result> RemoveRemoteDomainAsync(
        RemoveRemoteDomainRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.RemoveRemoteDomain, request, eventHandler, cancellationToken);

    public Task<Result<GetOrganizationRelationshipsResponse>> GetOrganizationRelationshipsAsync(
        GetOrganizationRelationshipsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetOrganizationRelationshipsResponse>(
            OperationType.GetOrganizationRelationships,
            request,
            eventHandler,
            cancellationToken);

    public Task<Result> UpsertOrganizationRelationshipAsync(
        UpsertOrganizationRelationshipRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.UpsertOrganizationRelationship, request, eventHandler, cancellationToken);

    public Task<Result> RemoveOrganizationRelationshipAsync(
        RemoveOrganizationRelationshipRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.RemoveOrganizationRelationship, request, eventHandler, cancellationToken);

    public Task<Result<GetAddressListsResponse>> GetAddressListsAsync(
        GetAddressListsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetAddressListsResponse>(OperationType.GetAddressLists, request, eventHandler, cancellationToken);

    public Task<Result> UpsertAddressListAsync(
        UpsertAddressListRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.UpsertAddressList, request, eventHandler, cancellationToken);

    public Task<Result> RemoveAddressListAsync(
        RemoveAddressListRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.RemoveAddressList, request, eventHandler, cancellationToken);

    public Task<Result<GetAddressBookPoliciesResponse>> GetAddressBookPoliciesAsync(
        GetAddressBookPoliciesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetAddressBookPoliciesResponse>(OperationType.GetAddressBookPolicies, request, eventHandler, cancellationToken);

    public Task<Result> UpsertAddressBookPolicyAsync(
        UpsertAddressBookPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.UpsertAddressBookPolicy, request, eventHandler, cancellationToken);

    public Task<Result> RemoveAddressBookPolicyAsync(
        RemoveAddressBookPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.RemoveAddressBookPolicy, request, eventHandler, cancellationToken);

    public Task<Result<GetOfflineAddressBooksResponse>> GetOfflineAddressBooksAsync(
        GetOfflineAddressBooksRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetOfflineAddressBooksResponse>(OperationType.GetOfflineAddressBooks, request, eventHandler, cancellationToken);

    public Task<Result> UpsertOfflineAddressBookAsync(
        UpsertOfflineAddressBookRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.UpsertOfflineAddressBook, request, eventHandler, cancellationToken);

    public Task<Result> RemoveOfflineAddressBookAsync(
        RemoveOfflineAddressBookRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.RemoveOfflineAddressBook, request, eventHandler, cancellationToken);

    public Task<Result<GetSharingPoliciesResponse>> GetSharingPoliciesAsync(
        GetSharingPoliciesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetSharingPoliciesResponse>(OperationType.GetSharingPolicies, request, eventHandler, cancellationToken);

    public Task<Result> UpsertSharingPolicyAsync(
        UpsertSharingPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.UpsertSharingPolicy, request, eventHandler, cancellationToken);

    public Task<Result> RemoveSharingPolicyAsync(
        RemoveSharingPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.RemoveSharingPolicy, request, eventHandler, cancellationToken);
}

