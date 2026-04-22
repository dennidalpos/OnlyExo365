using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Worker.PowerShell;

namespace OnlyExo365.Worker.Operations;

public partial class OperationDispatcher
{
    private async Task<ResponseEnvelope> HandleGetMessageTraceAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var traceRequest = JsonMessageSerializer.ExtractPayload<GetMessageTraceRequest>(request.Payload)
            ?? new GetMessageTraceRequest();

        await SendLogAsync(correlationId, LogLevel.Information, "Fetching message trace...");
        await SendProgressAsync(correlationId, 0, "Starting message trace query...");

        var response = await _exoCommands.GetMessageTraceAsync(
            traceRequest,
            cancellationToken);

        response.CorrelationId = correlationId;
        foreach (var warning in response.Warnings)
        {
            await SendLogAsync(correlationId, LogLevel.Warning, warning);
        }

        await SendProgressAsync(correlationId, 100, "Message trace complete");

        return CreateSuccessResponse(correlationId, response);
    }

    private async Task<ResponseEnvelope> HandleGetMessageTraceDetailsAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var detailsRequest = JsonMessageSerializer.ExtractPayload<GetMessageTraceDetailsRequest>(request.Payload);

        if (detailsRequest == null || string.IsNullOrWhiteSpace(detailsRequest.MessageTraceId) || string.IsNullOrWhiteSpace(detailsRequest.RecipientAddress))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "MessageTraceId and RecipientAddress are required");
        }

        await SendLogAsync(correlationId, LogLevel.Information, "Fetching message trace details...");
        await SendProgressAsync(correlationId, 0, "Starting details query...");

        var response = await _exoCommands.GetMessageTraceDetailsAsync(
            detailsRequest,
            cancellationToken);

        response.CorrelationId = correlationId;
        foreach (var warning in response.Warnings)
        {
            await SendLogAsync(correlationId, LogLevel.Warning, warning);
        }

        await SendProgressAsync(correlationId, 100, "Message trace details complete");

        return CreateSuccessResponse(correlationId, response);
    }

    private async Task<ResponseEnvelope> HandleGetTransportRulesAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        await SendLogAsync(correlationId, LogLevel.Information, "Fetching transport rules...");
        var response = await _exoCommands.GetTransportRulesAsync(cancellationToken);
        return CreateSuccessResponse(correlationId, response);
    }

    private async Task<ResponseEnvelope> HandleSetTransportRuleStateAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var stateRequest = JsonMessageSerializer.ExtractPayload<SetTransportRuleStateRequest>(request.Payload);
        if (stateRequest == null || string.IsNullOrWhiteSpace(stateRequest.Identity))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(correlationId, LogLevel.Information, $"Setting transport rule state: {stateRequest.Identity} => {(stateRequest.Enabled ? "Enabled" : "Disabled")}");
        await _exoCommands.SetTransportRuleStateAsync(stateRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleGetConnectorsAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        await SendLogAsync(correlationId, LogLevel.Information, "Fetching connectors...");
        var response = await _exoCommands.GetConnectorsAsync(cancellationToken);
        return CreateSuccessResponse(correlationId, response);
    }

    private async Task<ResponseEnvelope> HandleGetAcceptedDomainsAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        await SendLogAsync(correlationId, LogLevel.Information, "Fetching accepted domains...");
        var response = await _exoCommands.GetAcceptedDomainsAsync(cancellationToken);
        return CreateSuccessResponse(correlationId, response);
    }

    private async Task<ResponseEnvelope> HandleUpsertTransportRuleAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var upsertRequest = JsonMessageSerializer.ExtractPayload<UpsertTransportRuleRequest>(request.Payload);
        if (upsertRequest == null || string.IsNullOrWhiteSpace(upsertRequest.Name))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Name is required");
        }

        await SendLogAsync(correlationId, LogLevel.Information, $"Saving transport rule: {upsertRequest.Name}");
        await _exoCommands.UpsertTransportRuleAsync(upsertRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleRemoveTransportRuleAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var removeRequest = JsonMessageSerializer.ExtractPayload<RemoveTransportRuleRequest>(request.Payload);
        if (removeRequest == null || string.IsNullOrWhiteSpace(removeRequest.Identity))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(correlationId, LogLevel.Warning, $"Removing transport rule: {removeRequest.Identity}");
        await _exoCommands.RemoveTransportRuleAsync(removeRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleTestTransportRuleAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var testRequest = JsonMessageSerializer.ExtractPayload<TestTransportRuleRequest>(request.Payload);
        if (testRequest == null || string.IsNullOrWhiteSpace(testRequest.Sender) || string.IsNullOrWhiteSpace(testRequest.Recipient))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Sender and Recipient are required");
        }

        var response = await _exoCommands.TestTransportRuleAsync(testRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, response);
    }

    private async Task<ResponseEnvelope> HandleUpsertConnectorAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var upsertRequest = JsonMessageSerializer.ExtractPayload<UpsertConnectorRequest>(request.Payload);
        if (upsertRequest == null || string.IsNullOrWhiteSpace(upsertRequest.Name))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Name is required");
        }

        await SendLogAsync(correlationId, LogLevel.Information, $"Saving connector: {upsertRequest.Name} ({upsertRequest.Type})");
        await _exoCommands.UpsertConnectorAsync(upsertRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleRemoveConnectorAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var removeRequest = JsonMessageSerializer.ExtractPayload<RemoveConnectorRequest>(request.Payload);
        if (removeRequest == null || string.IsNullOrWhiteSpace(removeRequest.Identity) || string.IsNullOrWhiteSpace(removeRequest.Type))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Identity and Type are required");
        }

        await SendLogAsync(correlationId, LogLevel.Warning, $"Removing connector: {removeRequest.Identity} ({removeRequest.Type})");
        await _exoCommands.RemoveConnectorAsync(removeRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleUpsertAcceptedDomainAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var upsertRequest = JsonMessageSerializer.ExtractPayload<UpsertAcceptedDomainRequest>(request.Payload);
        if (upsertRequest == null || string.IsNullOrWhiteSpace(upsertRequest.Name) || string.IsNullOrWhiteSpace(upsertRequest.DomainName))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Name and DomainName are required");
        }

        await SendLogAsync(correlationId, LogLevel.Information, $"Saving accepted domain: {upsertRequest.DomainName}");
        await _exoCommands.UpsertAcceptedDomainAsync(upsertRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleRemoveAcceptedDomainAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var removeRequest = JsonMessageSerializer.ExtractPayload<RemoveAcceptedDomainRequest>(request.Payload);
        if (removeRequest == null || string.IsNullOrWhiteSpace(removeRequest.Identity))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(correlationId, LogLevel.Warning, $"Removing accepted domain: {removeRequest.Identity}");
        await _exoCommands.RemoveAcceptedDomainAsync(removeRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleGetRemoteDomainsAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        await SendLogAsync(correlationId, LogLevel.Information, "Fetching remote domains...");
        var response = await _exoCommands.GetRemoteDomainsAsync(cancellationToken);
        return CreateSuccessResponse(correlationId, response);
    }

    private async Task<ResponseEnvelope> HandleUpsertRemoteDomainAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var upsertRequest = JsonMessageSerializer.ExtractPayload<UpsertRemoteDomainRequest>(request.Payload);
        if (upsertRequest == null || string.IsNullOrWhiteSpace(upsertRequest.Name) || string.IsNullOrWhiteSpace(upsertRequest.DomainName))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Name and DomainName are required");
        }

        await SendLogAsync(correlationId, LogLevel.Information, $"Saving remote domain: {upsertRequest.DomainName}");
        await _exoCommands.UpsertRemoteDomainAsync(upsertRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleRemoveRemoteDomainAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var removeRequest = JsonMessageSerializer.ExtractPayload<RemoveRemoteDomainRequest>(request.Payload);
        if (removeRequest == null || string.IsNullOrWhiteSpace(removeRequest.Identity))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(correlationId, LogLevel.Warning, $"Removing remote domain: {removeRequest.Identity}");
        await _exoCommands.RemoveRemoteDomainAsync(removeRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleGetOrganizationRelationshipsAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        await SendLogAsync(correlationId, LogLevel.Information, "Fetching organization relationships...");
        var response = await _exoCommands.GetOrganizationRelationshipsAsync(cancellationToken);
        return CreateSuccessResponse(correlationId, response);
    }

    private async Task<ResponseEnvelope> HandleUpsertOrganizationRelationshipAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var upsertRequest = JsonMessageSerializer.ExtractPayload<UpsertOrganizationRelationshipRequest>(request.Payload);
        if (upsertRequest == null || string.IsNullOrWhiteSpace(upsertRequest.Name) || upsertRequest.DomainNames.Count == 0)
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Name and at least one DomainName are required");
        }

        await SendLogAsync(correlationId, LogLevel.Information, $"Saving organization relationship: {upsertRequest.Name}");
        await _exoCommands.UpsertOrganizationRelationshipAsync(upsertRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleRemoveOrganizationRelationshipAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var removeRequest = JsonMessageSerializer.ExtractPayload<RemoveOrganizationRelationshipRequest>(request.Payload);
        if (removeRequest == null || string.IsNullOrWhiteSpace(removeRequest.Identity))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(correlationId, LogLevel.Warning, $"Removing organization relationship: {removeRequest.Identity}");
        await _exoCommands.RemoveOrganizationRelationshipAsync(removeRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleGetAddressListsAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        await SendLogAsync(correlationId, LogLevel.Information, "Fetching address lists...");
        var response = await _exoCommands.GetAddressListsAsync(cancellationToken);
        return CreateSuccessResponse(correlationId, response);
    }

    private async Task<ResponseEnvelope> HandleUpsertAddressListAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var upsertRequest = JsonMessageSerializer.ExtractPayload<UpsertAddressListRequest>(request.Payload);
        if (upsertRequest == null || string.IsNullOrWhiteSpace(upsertRequest.Name))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Name is required");
        }

        await SendLogAsync(correlationId, LogLevel.Information, $"Saving address list: {upsertRequest.Name}");
        await _exoCommands.UpsertAddressListAsync(upsertRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleRemoveAddressListAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var removeRequest = JsonMessageSerializer.ExtractPayload<RemoveAddressListRequest>(request.Payload);
        if (removeRequest == null || string.IsNullOrWhiteSpace(removeRequest.Identity))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(correlationId, LogLevel.Warning, $"Removing address list: {removeRequest.Identity}");
        await _exoCommands.RemoveAddressListAsync(removeRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleGetAddressBookPoliciesAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        await SendLogAsync(correlationId, LogLevel.Information, "Fetching address book policies...");
        var response = await _exoCommands.GetAddressBookPoliciesAsync(cancellationToken);
        return CreateSuccessResponse(correlationId, response);
    }

    private async Task<ResponseEnvelope> HandleUpsertAddressBookPolicyAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var upsertRequest = JsonMessageSerializer.ExtractPayload<UpsertAddressBookPolicyRequest>(request.Payload);
        if (upsertRequest == null ||
            string.IsNullOrWhiteSpace(upsertRequest.Name) ||
            upsertRequest.AddressLists.Count == 0 ||
            string.IsNullOrWhiteSpace(upsertRequest.GlobalAddressList) ||
            string.IsNullOrWhiteSpace(upsertRequest.OfflineAddressBook))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Name, AddressLists, GlobalAddressList and OfflineAddressBook are required");
        }

        await SendLogAsync(correlationId, LogLevel.Information, $"Saving address book policy: {upsertRequest.Name}");
        await _exoCommands.UpsertAddressBookPolicyAsync(upsertRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleRemoveAddressBookPolicyAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var removeRequest = JsonMessageSerializer.ExtractPayload<RemoveAddressBookPolicyRequest>(request.Payload);
        if (removeRequest == null || string.IsNullOrWhiteSpace(removeRequest.Identity))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(correlationId, LogLevel.Warning, $"Removing address book policy: {removeRequest.Identity}");
        await _exoCommands.RemoveAddressBookPolicyAsync(removeRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleGetOfflineAddressBooksAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        await SendLogAsync(correlationId, LogLevel.Information, "Fetching offline address books...");
        var response = await _exoCommands.GetOfflineAddressBooksAsync(cancellationToken);
        return CreateSuccessResponse(correlationId, response);
    }

    private async Task<ResponseEnvelope> HandleUpsertOfflineAddressBookAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var upsertRequest = JsonMessageSerializer.ExtractPayload<UpsertOfflineAddressBookRequest>(request.Payload);
        if (upsertRequest == null || string.IsNullOrWhiteSpace(upsertRequest.Name) || upsertRequest.AddressLists.Count == 0)
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Name and AddressLists are required");
        }

        await SendLogAsync(correlationId, LogLevel.Information, $"Saving offline address book: {upsertRequest.Name}");
        await _exoCommands.UpsertOfflineAddressBookAsync(upsertRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleRemoveOfflineAddressBookAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var removeRequest = JsonMessageSerializer.ExtractPayload<RemoveOfflineAddressBookRequest>(request.Payload);
        if (removeRequest == null || string.IsNullOrWhiteSpace(removeRequest.Identity))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(correlationId, LogLevel.Warning, $"Removing offline address book: {removeRequest.Identity}");
        await _exoCommands.RemoveOfflineAddressBookAsync(removeRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleGetSharingPoliciesAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        await SendLogAsync(correlationId, LogLevel.Information, "Fetching sharing policies...");
        var response = await _exoCommands.GetSharingPoliciesAsync(cancellationToken);
        return CreateSuccessResponse(correlationId, response);
    }

    private async Task<ResponseEnvelope> HandleUpsertSharingPolicyAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var upsertRequest = JsonMessageSerializer.ExtractPayload<UpsertSharingPolicyRequest>(request.Payload);
        if (upsertRequest == null || string.IsNullOrWhiteSpace(upsertRequest.Name) || upsertRequest.Domains.Count == 0)
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Name and Domains are required");
        }

        await SendLogAsync(correlationId, LogLevel.Information, $"Saving sharing policy: {upsertRequest.Name}");
        await _exoCommands.UpsertSharingPolicyAsync(upsertRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleRemoveSharingPolicyAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var removeRequest = JsonMessageSerializer.ExtractPayload<RemoveSharingPolicyRequest>(request.Payload);
        if (removeRequest == null || string.IsNullOrWhiteSpace(removeRequest.Identity))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(correlationId, LogLevel.Warning, $"Removing sharing policy: {removeRequest.Identity}");
        await _exoCommands.RemoveSharingPolicyAsync(removeRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }
}

