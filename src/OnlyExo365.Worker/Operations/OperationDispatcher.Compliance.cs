using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;

namespace OnlyExo365.Worker.Operations;

public partial class OperationDispatcher
{
    private async Task<ResponseEnvelope> HandleGetComplianceWorkspaceAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var workspaceRequest = JsonMessageSerializer.ExtractPayload<GetComplianceWorkspaceRequest>(request.Payload)
            ?? new GetComplianceWorkspaceRequest();

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Loading compliance workspace...");
        var response = await _exoCommands.GetComplianceWorkspaceAsync(
            workspaceRequest,
            onLog: async (_, msg) => await SendLogAsync(request.CorrelationId, LogLevel.Verbose, msg),
            cancellationToken: cancellationToken);

        response.CorrelationId = request.CorrelationId;
        foreach (var warning in response.Warnings)
        {
            await SendLogAsync(request.CorrelationId, LogLevel.Warning, warning);
        }

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleSearchUnifiedAuditLogAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var auditRequest = JsonMessageSerializer.ExtractPayload<SearchUnifiedAuditLogRequest>(request.Payload);
        if (auditRequest == null)
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Audit search payload is required.");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Searching unified audit log...");
        await SendProgressAsync(request.CorrelationId, 0, "Running Search-UnifiedAuditLog...");

        var response = await _exoCommands.SearchUnifiedAuditLogAsync(
            auditRequest,
            onLog: async (_, msg) => await SendLogAsync(request.CorrelationId, LogLevel.Verbose, msg),
            cancellationToken: cancellationToken);

        await SendProgressAsync(request.CorrelationId, 100, "Unified audit log query complete");
        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleCreateComplianceSearchAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var createRequest = JsonMessageSerializer.ExtractPayload<CreateComplianceSearchRequest>(request.Payload);
        if (createRequest == null || string.IsNullOrWhiteSpace(createRequest.Name))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Name is required.");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Creating compliance search: {createRequest.Name}");
        await _exoCommands.CreateComplianceSearchAsync(
            createRequest,
            onLog: async (_, msg) => await SendLogAsync(request.CorrelationId, LogLevel.Verbose, msg),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleStartComplianceSearchAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var startRequest = JsonMessageSerializer.ExtractPayload<StartComplianceSearchRequest>(request.Payload);
        if (startRequest == null || string.IsNullOrWhiteSpace(startRequest.Name))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Name is required.");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Starting compliance search: {startRequest.Name}");
        await _exoCommands.StartComplianceSearchAsync(
            startRequest,
            onLog: async (_, msg) => await SendLogAsync(request.CorrelationId, LogLevel.Verbose, msg),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleRemoveComplianceSearchAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var removeRequest = JsonMessageSerializer.ExtractPayload<RemoveComplianceSearchRequest>(request.Payload);
        if (removeRequest == null || string.IsNullOrWhiteSpace(removeRequest.Name))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Name is required.");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Warning, $"Removing compliance search: {removeRequest.Name}");
        await _exoCommands.RemoveComplianceSearchAsync(
            removeRequest,
            onLog: async (_, msg) => await SendLogAsync(request.CorrelationId, LogLevel.Verbose, msg),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleInvokeComplianceActionAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var actionRequest = JsonMessageSerializer.ExtractPayload<InvokeComplianceActionRequest>(request.Payload);
        if (actionRequest == null || string.IsNullOrWhiteSpace(actionRequest.SearchName) || string.IsNullOrWhiteSpace(actionRequest.ActionType))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "SearchName and ActionType are required.");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Invoking compliance action {actionRequest.ActionType} on {actionRequest.SearchName}");
        var response = await _exoCommands.InvokeComplianceActionAsync(
            actionRequest,
            onLog: async (_, msg) => await SendLogAsync(request.CorrelationId, LogLevel.Verbose, msg),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, response);
    }
}

