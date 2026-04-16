using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Worker.PowerShell;

namespace ExchangeAdmin.Worker.Operations;

public partial class OperationDispatcher
{
    private async Task<ResponseEnvelope> HandleGetMailboxProvisioningCandidatesAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var provisioningRequest = JsonMessageSerializer.ExtractPayload<GetMailboxProvisioningCandidatesRequest>(request.Payload)
            ?? new GetMailboxProvisioningCandidatesRequest();

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Fetching member users for mailbox provisioning...");
        await SendProgressAsync(request.CorrelationId, 0, "Fetching member users...");

        var response = await _exoCommands.GetMailboxProvisioningCandidatesAsync(provisioningRequest, cancellationToken);

        await SendProgressAsync(
            request.CorrelationId,
            100,
            FormatListProgressStatus(response.Candidates.Count, response.TotalCount, "Analizzati"),
            response.Candidates.Count,
            response.TotalCount);

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleGetMailboxesAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var mailboxRequest = JsonMessageSerializer.ExtractPayload<GetMailboxesRequest>(request.Payload)
            ?? new GetMailboxesRequest();

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Fetching mailboxes...");
        await SendProgressAsync(request.CorrelationId, 0, "Starting mailbox retrieval...");

        var response = await _exoCommands.GetMailboxesAsync(
            mailboxRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            onPartialOutput: async (item) => await SendPartialOutputAsync(request.CorrelationId, item, 0),
            cancellationToken: cancellationToken);

        await SendProgressAsync(
            request.CorrelationId,
            100,
            FormatListProgressStatus(response.Mailboxes.Count, response.TotalCount, "Analizzate"),
            response.Mailboxes.Count,
            response.TotalCount);

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleGetDeletedMailboxesAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var mailboxRequest = JsonMessageSerializer.ExtractPayload<GetDeletedMailboxesRequest>(request.Payload)
            ?? new GetDeletedMailboxesRequest();

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Fetching deleted mailboxes...");
        await SendProgressAsync(request.CorrelationId, 0, "Starting deleted mailbox retrieval...");

        var response = await _exoCommands.GetDeletedMailboxesAsync(
            mailboxRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        await SendProgressAsync(
            request.CorrelationId,
            100,
            FormatListProgressStatus(response.Mailboxes.Count, response.TotalCount, "Analizzate"),
            response.Mailboxes.Count,
            response.TotalCount);

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleGetMailboxDetailsAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var detailsRequest = JsonMessageSerializer.ExtractPayload<GetMailboxDetailsRequest>(request.Payload);

        if (detailsRequest == null || string.IsNullOrWhiteSpace(detailsRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Fetching mailbox details for {detailsRequest.Identity}...");
        await SendProgressAsync(request.CorrelationId, 0, "Fetching mailbox details...");

        var details = await _exoCommands.GetMailboxDetailsAsync(
            detailsRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        await SendProgressAsync(request.CorrelationId, 100, "Mailbox details retrieved");

        return CreateSuccessResponse(request.CorrelationId, details);
    }

    private async Task<ResponseEnvelope> HandleGetRetentionPoliciesAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        _ = JsonMessageSerializer.ExtractPayload<GetRetentionPoliciesRequest>(request.Payload)
            ?? new GetRetentionPoliciesRequest();

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Fetching retention policies...");
        await SendProgressAsync(request.CorrelationId, 0, "Fetching retention policies...");

        var policies = await _exoCommands.GetRetentionPoliciesAsync(
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        var response = new GetRetentionPoliciesResponse
        {
            Policies = policies
        };

        await SendProgressAsync(request.CorrelationId, 100, "Retention policies retrieved");

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleSetRetentionPolicyAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var setRequest = JsonMessageSerializer.ExtractPayload<SetRetentionPolicyRequest>(request.Payload);

        if (setRequest == null || string.IsNullOrWhiteSpace(setRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendProgressAsync(request.CorrelationId, 0, "Setting retention policy...");

        await _exoCommands.SetRetentionPolicyAsync(
            setRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        await SendProgressAsync(request.CorrelationId, 100, "Retention policy applied");

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleGetMailboxPermissionsAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var permRequest = JsonMessageSerializer.ExtractPayload<GetMailboxPermissionsRequest>(request.Payload);

        if (permRequest == null || string.IsNullOrWhiteSpace(permRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Fetching permissions for {permRequest.Identity}...");
        await SendProgressAsync(request.CorrelationId, 0, "Fetching mailbox permissions...");

        var permissions = await _exoCommands.GetMailboxPermissionsAsync(
            permRequest.Identity,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        await SendProgressAsync(request.CorrelationId, 100, "Mailbox permissions retrieved");

        return CreateSuccessResponse(request.CorrelationId, permissions);
    }

    private async Task<ResponseEnvelope> HandleGetMailboxFolderPermissionsAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var folderRequest = JsonMessageSerializer.ExtractPayload<GetMailboxFolderPermissionsRequest>(request.Payload);

        if (folderRequest == null ||
            string.IsNullOrWhiteSpace(folderRequest.MailboxIdentity) ||
            string.IsNullOrWhiteSpace(folderRequest.FolderPath))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "MailboxIdentity and FolderPath are required");
        }

        await SendLogAsync(
            request.CorrelationId,
            LogLevel.Information,
            $"Fetching folder permissions for {folderRequest.MailboxIdentity}:{folderRequest.FolderPath}...");
        await SendProgressAsync(request.CorrelationId, 0, "Fetching mailbox folder permissions...");

        var response = await _exoCommands.GetMailboxFolderPermissionsAsync(
            folderRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        await SendProgressAsync(request.CorrelationId, 100, "Mailbox folder permissions retrieved");

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleSetMailboxPermissionAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var setRequest = JsonMessageSerializer.ExtractPayload<SetMailboxPermissionRequest>(request.Payload);

        if (setRequest == null || string.IsNullOrWhiteSpace(setRequest.Identity) || string.IsNullOrWhiteSpace(setRequest.User))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity and User are required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information,
            $"{setRequest.Action} {setRequest.PermissionType} for {setRequest.User} on {setRequest.Identity}...");
        await SendProgressAsync(request.CorrelationId, 0, "Setting mailbox permission...");

        await _exoCommands.SetMailboxPermissionAsync(
            setRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        await SendProgressAsync(request.CorrelationId, 100, "Mailbox permission applied");

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleSetMailboxFolderPermissionAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var setRequest = JsonMessageSerializer.ExtractPayload<SetMailboxFolderPermissionRequest>(request.Payload);

        if (setRequest == null ||
            string.IsNullOrWhiteSpace(setRequest.MailboxIdentity) ||
            string.IsNullOrWhiteSpace(setRequest.FolderPath) ||
            string.IsNullOrWhiteSpace(setRequest.User))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "MailboxIdentity, FolderPath and User are required");
        }

        if (setRequest.Action != PermissionAction.Remove && setRequest.AccessRights.Count == 0)
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "AccessRights are required for add or modify actions");
        }

        await SendLogAsync(
            request.CorrelationId,
            LogLevel.Information,
            $"{setRequest.Action} folder permission for {setRequest.User} on {setRequest.MailboxIdentity}:{setRequest.FolderPath}...");
        await SendProgressAsync(request.CorrelationId, 0, "Applying mailbox folder permission...");

        await _exoCommands.SetMailboxFolderPermissionAsync(
            setRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        await SendProgressAsync(request.CorrelationId, 100, "Mailbox folder permission applied");

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleApplyPermissionsDeltaPlanAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var applyRequest = JsonMessageSerializer.ExtractPayload<ApplyPermissionsDeltaPlanRequest>(request.Payload);

        if (applyRequest == null || string.IsNullOrWhiteSpace(applyRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information,
            $"Applying permissions delta plan ({applyRequest.Actions.Count} actions)...");

        var response = await _exoCommands.ApplyPermissionsDeltaPlanAsync(
            applyRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            onProgress: async (current, total) =>
            {
                var percent = (int)(current * 100.0 / total);
                await SendProgressAsync(request.CorrelationId, percent, $"Applying action {current} of {total}", current, total);
            },
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleUpdateMailboxSettingsAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var settingsRequest = JsonMessageSerializer.ExtractPayload<UpdateMailboxSettingsRequest>(request.Payload);

        if (settingsRequest == null || string.IsNullOrWhiteSpace(settingsRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendProgressAsync(request.CorrelationId, 0, "Updating mailbox settings...");

        await _exoCommands.SetMailboxSettingsAsync(
            settingsRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        await SendProgressAsync(request.CorrelationId, 100, "Mailbox settings updated");

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleSetMailboxAutoReplyConfigurationAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var autoReplyRequest = JsonMessageSerializer.ExtractPayload<SetMailboxAutoReplyConfigurationRequest>(request.Payload);

        if (autoReplyRequest == null || string.IsNullOrWhiteSpace(autoReplyRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendProgressAsync(request.CorrelationId, 0, "Setting auto-reply configuration...");

        await _exoCommands.SetMailboxAutoReplyConfigurationAsync(
            autoReplyRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        await SendProgressAsync(request.CorrelationId, 100, "Auto-reply configuration updated");

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleCreateMailboxAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var createRequest = JsonMessageSerializer.ExtractPayload<CreateMailboxRequest>(request.Payload);
        if (createRequest != null)
        {
            WorkerSecretResolver.Resolve(createRequest);
        }

        if (createRequest == null)
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Missing create mailbox request payload", false, null);
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Creating mailbox {createRequest.PrimarySmtpAddress}...");

        await _exoCommands.CreateMailboxAsync(
            createRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, LogLevel.Verbose, msg),
            cancellationToken: cancellationToken);

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Mailbox created successfully");

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleConvertMailboxToSharedAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var convertRequest = JsonMessageSerializer.ExtractPayload<ConvertMailboxToSharedRequest>(request.Payload);

        if (convertRequest == null || string.IsNullOrWhiteSpace(convertRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendProgressAsync(request.CorrelationId, 0, "Converting mailbox to shared...");

        await _exoCommands.ConvertMailboxToSharedAsync(
            convertRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        await SendProgressAsync(request.CorrelationId, 100, "Mailbox converted to shared");

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleConvertMailboxToRegularAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var convertRequest = JsonMessageSerializer.ExtractPayload<ConvertMailboxToRegularRequest>(request.Payload);

        if (convertRequest == null || string.IsNullOrWhiteSpace(convertRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendProgressAsync(request.CorrelationId, 0, "Converting mailbox to regular...");

        await _exoCommands.ConvertMailboxToRegularAsync(
            convertRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        await SendProgressAsync(request.CorrelationId, 100, "Mailbox converted to regular");

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleRestoreMailboxAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var restoreRequest = JsonMessageSerializer.ExtractPayload<RestoreMailboxRequest>(request.Payload);

        if (restoreRequest == null || string.IsNullOrWhiteSpace(restoreRequest.SourceIdentity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "SourceIdentity is required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Starting mailbox restore for {restoreRequest.SourceIdentity}...");
        await SendProgressAsync(request.CorrelationId, 0, "Detecting mailbox state...");

        var response = await _exoCommands.RestoreMailboxAsync(
            restoreRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        if (response.Status == RestoreMailboxStatus.InProgress && response.PercentComplete.HasValue)
        {
            await SendProgressAsync(request.CorrelationId, response.PercentComplete.Value, "Mailbox restore in progress...");
        }
        else if (response.Status == RestoreMailboxStatus.Completed)
        {
            await SendProgressAsync(request.CorrelationId, 100, "Mailbox restore completed");
        }
        else if (response.Status == RestoreMailboxStatus.Failed)
        {
            await SendProgressAsync(request.CorrelationId, 100, "Mailbox restore failed");
        }
        else
        {
            await SendProgressAsync(request.CorrelationId, 100, "Mailbox restore request submitted");
        }

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleGetMailboxSpaceReportAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var reportRequest = JsonMessageSerializer.ExtractPayload<GetMailboxSpaceReportRequest>(request.Payload)
            ?? new GetMailboxSpaceReportRequest();

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Fetching mailbox space report...");
        await SendProgressAsync(request.CorrelationId, 0, "Fetching mailbox space report...");

        var lastPercent = -1;

        var response = await _exoCommands.GetMailboxSpaceReportAsync(
            reportRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            onProgress: (current, total) =>
            {
                var percent = total > 0 ? (int)Math.Round(current * 100.0 / total) : 0;
                if (percent == lastPercent && current != total)
                {
                    return;
                }

                lastPercent = percent;
                var remaining = Math.Max(0, total - current);
                var status = $"Analizzate {current}/{total} (rimanenti {remaining})";
                _ = SendProgressAsync(request.CorrelationId, percent, status, current, total);
            },
            cancellationToken: cancellationToken);

        response.CorrelationId = request.CorrelationId;
        foreach (var warning in response.Warnings)
        {
            await SendLogAsync(request.CorrelationId, LogLevel.Warning, warning);
        }

        await SendProgressAsync(request.CorrelationId, 100, "Mailbox space report complete");

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleGetMailboxAccessReportAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var reportRequest = JsonMessageSerializer.ExtractPayload<GetMailboxAccessReportRequest>(request.Payload)
            ?? new GetMailboxAccessReportRequest();

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Fetching mailbox access report...");
        await SendProgressAsync(request.CorrelationId, 0, "Fetching mailbox access report...");

        var lastPercent = -1;

        var response = await _exoCommands.GetMailboxAccessReportAsync(
            reportRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            onProgress: (current, total) =>
            {
                var percent = total > 0 ? (int)Math.Round(current * 100.0 / total) : 0;
                if (percent == lastPercent && current != total)
                {
                    return;
                }

                lastPercent = percent;
                var remaining = Math.Max(0, total - current);
                var status = $"Analizzate {current}/{total} mailbox (rimanenti {remaining})";
                _ = SendProgressAsync(request.CorrelationId, percent, status, current, total);
            },
            cancellationToken: cancellationToken);

        response.CorrelationId = request.CorrelationId;
        foreach (var warning in response.Warnings)
        {
            await SendLogAsync(request.CorrelationId, LogLevel.Warning, warning);
        }

        await SendProgressAsync(request.CorrelationId, 100, "Mailbox access report complete");

        return CreateSuccessResponse(request.CorrelationId, response);
    }
}
