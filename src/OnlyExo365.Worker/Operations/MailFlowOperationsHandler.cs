using OnlyExo365.Contracts.Messages;

namespace OnlyExo365.Worker.Operations;

public partial class OperationDispatcher
{
    private sealed class MailFlowOperationsHandler(OperationDispatcher dispatcher) : IOperationAreaHandler
    {
        public IReadOnlyCollection<OperationType> SupportedOperations { get; } =
        [
            OperationType.GetMessageTrace,
            OperationType.GetMessageTraceDetails,
            OperationType.GetTransportRules,
            OperationType.SetTransportRuleState,
            OperationType.UpsertTransportRule,
            OperationType.RemoveTransportRule,
            OperationType.TestTransportRule,
            OperationType.GetConnectors,
            OperationType.UpsertConnector,
            OperationType.RemoveConnector,
            OperationType.GetAcceptedDomains,
            OperationType.UpsertAcceptedDomain,
            OperationType.RemoveAcceptedDomain,
            OperationType.GetRemoteDomains,
            OperationType.UpsertRemoteDomain,
            OperationType.RemoveRemoteDomain,
            OperationType.GetOrganizationRelationships,
            OperationType.UpsertOrganizationRelationship,
            OperationType.RemoveOrganizationRelationship,
            OperationType.GetAddressLists,
            OperationType.UpsertAddressList,
            OperationType.RemoveAddressList,
            OperationType.GetAddressBookPolicies,
            OperationType.UpsertAddressBookPolicy,
            OperationType.RemoveAddressBookPolicy,
            OperationType.GetOfflineAddressBooks,
            OperationType.UpsertOfflineAddressBook,
            OperationType.RemoveOfflineAddressBook,
            OperationType.GetSharingPolicies,
            OperationType.UpsertSharingPolicy,
            OperationType.RemoveSharingPolicy
        ];

        public Task<ResponseEnvelope> HandleAsync(RequestEnvelope request, CancellationToken cancellationToken)
        {
            var correlationId = request.CorrelationId;

            return request.Operation switch
            {
                OperationType.GetMessageTrace => dispatcher.HandleGetMessageTraceAsync(request, correlationId, cancellationToken),
                OperationType.GetMessageTraceDetails => dispatcher.HandleGetMessageTraceDetailsAsync(request, correlationId, cancellationToken),
                OperationType.GetTransportRules => dispatcher.HandleGetTransportRulesAsync(request, correlationId, cancellationToken),
                OperationType.SetTransportRuleState => dispatcher.HandleSetTransportRuleStateAsync(request, correlationId, cancellationToken),
                OperationType.UpsertTransportRule => dispatcher.HandleUpsertTransportRuleAsync(request, correlationId, cancellationToken),
                OperationType.RemoveTransportRule => dispatcher.HandleRemoveTransportRuleAsync(request, correlationId, cancellationToken),
                OperationType.TestTransportRule => dispatcher.HandleTestTransportRuleAsync(request, correlationId, cancellationToken),
                OperationType.GetConnectors => dispatcher.HandleGetConnectorsAsync(request, correlationId, cancellationToken),
                OperationType.UpsertConnector => dispatcher.HandleUpsertConnectorAsync(request, correlationId, cancellationToken),
                OperationType.RemoveConnector => dispatcher.HandleRemoveConnectorAsync(request, correlationId, cancellationToken),
                OperationType.GetAcceptedDomains => dispatcher.HandleGetAcceptedDomainsAsync(request, correlationId, cancellationToken),
                OperationType.UpsertAcceptedDomain => dispatcher.HandleUpsertAcceptedDomainAsync(request, correlationId, cancellationToken),
                OperationType.RemoveAcceptedDomain => dispatcher.HandleRemoveAcceptedDomainAsync(request, correlationId, cancellationToken),
                OperationType.GetRemoteDomains => dispatcher.HandleGetRemoteDomainsAsync(request, correlationId, cancellationToken),
                OperationType.UpsertRemoteDomain => dispatcher.HandleUpsertRemoteDomainAsync(request, correlationId, cancellationToken),
                OperationType.RemoveRemoteDomain => dispatcher.HandleRemoveRemoteDomainAsync(request, correlationId, cancellationToken),
                OperationType.GetOrganizationRelationships => dispatcher.HandleGetOrganizationRelationshipsAsync(request, correlationId, cancellationToken),
                OperationType.UpsertOrganizationRelationship => dispatcher.HandleUpsertOrganizationRelationshipAsync(request, correlationId, cancellationToken),
                OperationType.RemoveOrganizationRelationship => dispatcher.HandleRemoveOrganizationRelationshipAsync(request, correlationId, cancellationToken),
                OperationType.GetAddressLists => dispatcher.HandleGetAddressListsAsync(request, correlationId, cancellationToken),
                OperationType.UpsertAddressList => dispatcher.HandleUpsertAddressListAsync(request, correlationId, cancellationToken),
                OperationType.RemoveAddressList => dispatcher.HandleRemoveAddressListAsync(request, correlationId, cancellationToken),
                OperationType.GetAddressBookPolicies => dispatcher.HandleGetAddressBookPoliciesAsync(request, correlationId, cancellationToken),
                OperationType.UpsertAddressBookPolicy => dispatcher.HandleUpsertAddressBookPolicyAsync(request, correlationId, cancellationToken),
                OperationType.RemoveAddressBookPolicy => dispatcher.HandleRemoveAddressBookPolicyAsync(request, correlationId, cancellationToken),
                OperationType.GetOfflineAddressBooks => dispatcher.HandleGetOfflineAddressBooksAsync(request, correlationId, cancellationToken),
                OperationType.UpsertOfflineAddressBook => dispatcher.HandleUpsertOfflineAddressBookAsync(request, correlationId, cancellationToken),
                OperationType.RemoveOfflineAddressBook => dispatcher.HandleRemoveOfflineAddressBookAsync(request, correlationId, cancellationToken),
                OperationType.GetSharingPolicies => dispatcher.HandleGetSharingPoliciesAsync(request, correlationId, cancellationToken),
                OperationType.UpsertSharingPolicy => dispatcher.HandleUpsertSharingPolicyAsync(request, correlationId, cancellationToken),
                OperationType.RemoveSharingPolicy => dispatcher.HandleRemoveSharingPolicyAsync(request, correlationId, cancellationToken),
                _ => throw new InvalidOperationException($"Unsupported mail flow operation: {request.Operation}")
            };
        }
    }
}

