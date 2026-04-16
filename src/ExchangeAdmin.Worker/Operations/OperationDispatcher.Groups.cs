using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Worker.PowerShell;

namespace ExchangeAdmin.Worker.Operations;

public partial class OperationDispatcher
{
    private async Task<ResponseEnvelope> HandleGetDistributionListsAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var listRequest = JsonMessageSerializer.ExtractPayload<GetDistributionListsRequest>(request.Payload)
            ?? new GetDistributionListsRequest();

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Fetching distribution lists...");
        await SendProgressAsync(request.CorrelationId, 0, "Fetching distribution lists...");

        var response = await _exoGroupCommands.GetDistributionListsAsync(
            listRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            onPartialOutput: async (item) => await SendPartialOutputAsync(request.CorrelationId, item, 0),
            cancellationToken: cancellationToken);

        await SendProgressAsync(
            request.CorrelationId,
            100,
            FormatListProgressStatus(response.DistributionLists.Count, response.TotalCount, "Analizzati"),
            response.DistributionLists.Count,
            response.TotalCount);

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleGetDistributionListDetailsAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var detailsRequest = JsonMessageSerializer.ExtractPayload<GetDistributionListDetailsRequest>(request.Payload);

        if (detailsRequest == null || string.IsNullOrWhiteSpace(detailsRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Fetching distribution list details for {detailsRequest.Identity}...");
        await SendProgressAsync(request.CorrelationId, 0, "Fetching distribution list details...");

        var details = await _exoGroupCommands.GetDistributionListDetailsAsync(
            detailsRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        await SendProgressAsync(request.CorrelationId, 100, "Distribution list details retrieved");

        return CreateSuccessResponse(request.CorrelationId, details);
    }

    private async Task<ResponseEnvelope> HandleGetGroupMembersAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var membersRequest = JsonMessageSerializer.ExtractPayload<GetGroupMembersRequest>(request.Payload);

        if (membersRequest == null || string.IsNullOrWhiteSpace(membersRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Fetching members for {membersRequest.Identity}...");
        await SendProgressAsync(request.CorrelationId, 0, "Fetching group members...");

        var members = await _exoGroupCommands.GetGroupMembersPageAsync(
            membersRequest.Identity,
            membersRequest.GroupType,
            membersRequest.Skip,
            membersRequest.PageSize,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        await SendProgressAsync(request.CorrelationId, 100, "Group members retrieved");

        return CreateSuccessResponse(request.CorrelationId, members);
    }

    private async Task<ResponseEnvelope> HandleModifyGroupMemberAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var modifyRequest = JsonMessageSerializer.ExtractPayload<ModifyGroupMemberRequest>(request.Payload);

        if (modifyRequest == null || string.IsNullOrWhiteSpace(modifyRequest.Identity) || string.IsNullOrWhiteSpace(modifyRequest.Member))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity and Member are required");
        }

        await SendProgressAsync(request.CorrelationId, 0, "Modifying group member...");

        await _exoGroupCommands.ModifyGroupMemberAsync(
            modifyRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        await SendProgressAsync(request.CorrelationId, 100, "Group member modified");

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandlePreviewDynamicGroupMembersAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var previewRequest = JsonMessageSerializer.ExtractPayload<PreviewDynamicGroupMembersRequest>(request.Payload);

        if (previewRequest == null || string.IsNullOrWhiteSpace(previewRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Warning, "Previewing dynamic group members (this may take a while)...");
        await SendProgressAsync(request.CorrelationId, 0, "Previewing dynamic group members...");

        var response = await _exoGroupCommands.PreviewDynamicGroupMembersAsync(
            previewRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        await SendProgressAsync(request.CorrelationId, 100, "Dynamic group preview complete");

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleCreateDistributionListAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var createRequest = JsonMessageSerializer.ExtractPayload<CreateDistributionListRequest>(request.Payload);
        if (createRequest == null)
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Missing create distribution list request payload", false, null);
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Creating distribution list {createRequest.PrimarySmtpAddress}...");

        await _exoGroupCommands.CreateDistributionListAsync(
            createRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, LogLevel.Verbose, msg),
            cancellationToken: cancellationToken);

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Distribution list created successfully");

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleSetDistributionListSettingsAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var settingsRequest = JsonMessageSerializer.ExtractPayload<SetDistributionListSettingsRequest>(request.Payload);

        if (settingsRequest == null || string.IsNullOrWhiteSpace(settingsRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendProgressAsync(request.CorrelationId, 0, "Updating distribution list settings...");

        await _exoGroupCommands.SetDistributionListSettingsAsync(
            settingsRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        await SendProgressAsync(request.CorrelationId, 100, "Distribution list settings updated");

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }
}
