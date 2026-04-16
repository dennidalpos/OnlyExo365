using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;

namespace ExchangeAdmin.Worker.Operations;

public partial class OperationDispatcher
{
    private async Task<ResponseEnvelope> HandleGetMailSecurityBaselineAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var baselineRequest = JsonMessageSerializer.ExtractPayload<GetMailSecurityBaselineRequest>(request.Payload)
            ?? new GetMailSecurityBaselineRequest();

        await SendLogAsync(correlationId, LogLevel.Information, "Fetching mail security baseline...");
        var response = await _exoCommands.GetMailSecurityBaselineAsync(baselineRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, response);
    }

    private async Task<ResponseEnvelope> HandleUpdateDkimSigningConfigAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var updateRequest = JsonMessageSerializer.ExtractPayload<UpdateDkimSigningConfigRequest>(request.Payload);
        if (updateRequest == null || string.IsNullOrWhiteSpace(updateRequest.Identity))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(correlationId, LogLevel.Information, $"Updating DKIM signing config: {updateRequest.Identity}");
        await _exoCommands.UpdateDkimSigningConfigAsync(updateRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleUpdateHostedContentFilterPolicyAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var updateRequest = JsonMessageSerializer.ExtractPayload<UpdateHostedContentFilterPolicyRequest>(request.Payload);
        if (updateRequest == null || string.IsNullOrWhiteSpace(updateRequest.Identity))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(correlationId, LogLevel.Information, $"Updating anti-spam policy: {updateRequest.Identity}");
        await _exoCommands.UpdateHostedContentFilterPolicyAsync(updateRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleUpdateAntiPhishPolicyAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var updateRequest = JsonMessageSerializer.ExtractPayload<UpdateAntiPhishPolicyRequest>(request.Payload);
        if (updateRequest == null || string.IsNullOrWhiteSpace(updateRequest.Identity))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(correlationId, LogLevel.Information, $"Updating anti-phish policy: {updateRequest.Identity}");
        await _exoCommands.UpdateAntiPhishPolicyAsync(updateRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleUpdateMalwareFilterPolicyAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var updateRequest = JsonMessageSerializer.ExtractPayload<UpdateMalwareFilterPolicyRequest>(request.Payload);
        if (updateRequest == null || string.IsNullOrWhiteSpace(updateRequest.Identity))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(correlationId, LogLevel.Information, $"Updating anti-malware policy: {updateRequest.Identity}");
        await _exoCommands.UpdateMalwareFilterPolicyAsync(updateRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleUpdateQuarantinePolicyAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var updateRequest = JsonMessageSerializer.ExtractPayload<UpdateQuarantinePolicyRequest>(request.Payload);
        if (updateRequest == null || string.IsNullOrWhiteSpace(updateRequest.Identity))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(correlationId, LogLevel.Information, $"Updating quarantine policy: {updateRequest.Identity}");
        await _exoCommands.UpdateQuarantinePolicyAsync(updateRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleUpdateHostedOutboundSpamFilterPolicyAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var updateRequest = JsonMessageSerializer.ExtractPayload<UpdateHostedOutboundSpamFilterPolicyRequest>(request.Payload);
        if (updateRequest == null || string.IsNullOrWhiteSpace(updateRequest.Identity))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(correlationId, LogLevel.Information, $"Updating outbound spam policy: {updateRequest.Identity}");
        await _exoCommands.UpdateHostedOutboundSpamFilterPolicyAsync(updateRequest, cancellationToken);
        return CreateSuccessResponse(correlationId, new { Success = true });
    }
}
