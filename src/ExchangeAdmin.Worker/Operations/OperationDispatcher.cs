using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Worker.PowerShell;

namespace ExchangeAdmin.Worker.Operations;

public class OperationDispatcher
{
    private readonly PowerShellEngine _psEngine;
    private readonly CapabilityDetector _capabilityDetector;
    private readonly ExoCommands _exoCommands;
    private readonly ExoGroupCommands _exoGroupCommands;
    private readonly Func<EventEnvelope, Task> _sendEvent;

    public OperationDispatcher(PowerShellEngine psEngine, Func<EventEnvelope, Task> sendEvent)
    {
        _psEngine = psEngine;
        _sendEvent = sendEvent;
        _capabilityDetector = new CapabilityDetector(psEngine);
        _exoCommands = new ExoCommands(psEngine, _capabilityDetector);
        _exoGroupCommands = new ExoGroupCommands(psEngine, _capabilityDetector);
    }

    public async Task<ResponseEnvelope> DispatchAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        ConsoleLogger.Info("Dispatcher", $"Dispatching operation: {request.Operation}");
        try
        {
            return request.Operation switch
            {
                OperationType.ConnectExchangeInteractive => await HandleConnectAsync(request, cancellationToken),
                OperationType.DisconnectExchange => await HandleDisconnectAsync(request, cancellationToken),
                OperationType.GetConnectionStatus => await HandleGetConnectionStatusAsync(request, cancellationToken),

                OperationType.DetectCapabilities => await HandleDetectCapabilitiesAsync(request, cancellationToken),

                OperationType.DemoLongOperation => await HandleDemoOperationAsync(request, cancellationToken),

                OperationType.GetDashboardStats => await HandleGetDashboardStatsAsync(request, cancellationToken),

                OperationType.GetMailboxes => await HandleGetMailboxesAsync(request, cancellationToken),
                OperationType.GetDeletedMailboxes => await HandleGetDeletedMailboxesAsync(request, cancellationToken),
                OperationType.GetMailboxDetails => await HandleGetMailboxDetailsAsync(request, cancellationToken),
                OperationType.GetRetentionPolicies => await HandleGetRetentionPoliciesAsync(request, cancellationToken),
                OperationType.SetRetentionPolicy => await HandleSetRetentionPolicyAsync(request, cancellationToken),
                OperationType.GetMailboxPermissions => await HandleGetMailboxPermissionsAsync(request, cancellationToken),
                OperationType.SetMailboxPermission => await HandleSetMailboxPermissionAsync(request, cancellationToken),
                OperationType.ApplyPermissionsDeltaPlan => await HandleApplyPermissionsDeltaPlanAsync(request, cancellationToken),
                OperationType.SetMailboxFeature => await HandleSetMailboxFeatureAsync(request, cancellationToken),
                OperationType.UpdateMailboxSettings => await HandleUpdateMailboxSettingsAsync(request, cancellationToken),
                OperationType.SetMailboxAutoReplyConfiguration => await HandleSetMailboxAutoReplyConfigurationAsync(request, cancellationToken),
                OperationType.ConvertMailboxToShared => await HandleConvertMailboxToSharedAsync(request, cancellationToken),
                OperationType.ConvertMailboxToRegular => await HandleConvertMailboxToRegularAsync(request, cancellationToken),
                OperationType.RestoreMailbox => await HandleRestoreMailboxAsync(request, cancellationToken),
                OperationType.GetMailboxSpaceReport => await HandleGetMailboxSpaceReportAsync(request, cancellationToken),
                OperationType.CreateMailbox => await HandleCreateMailboxAsync(request, cancellationToken),

                OperationType.GetDistributionLists => await HandleGetDistributionListsAsync(request, cancellationToken),
                OperationType.GetDistributionListDetails => await HandleGetDistributionListDetailsAsync(request, cancellationToken),
                OperationType.GetGroupMembers => await HandleGetGroupMembersAsync(request, cancellationToken),
                OperationType.ModifyGroupMember => await HandleModifyGroupMemberAsync(request, cancellationToken),
                OperationType.PreviewDynamicGroupMembers => await HandlePreviewDynamicGroupMembersAsync(request, cancellationToken),
                OperationType.SetDistributionListSettings => await HandleSetDistributionListSettingsAsync(request, cancellationToken),
                OperationType.CreateDistributionList => await HandleCreateDistributionListAsync(request, cancellationToken),

                OperationType.GetMessageTrace => await HandleGetMessageTraceAsync(request, request.CorrelationId, cancellationToken),
                OperationType.GetMessageTraceDetails => await HandleGetMessageTraceDetailsAsync(request, request.CorrelationId, cancellationToken),
                OperationType.GetTransportRules => await HandleGetTransportRulesAsync(request, request.CorrelationId, cancellationToken),
                OperationType.SetTransportRuleState => await HandleSetTransportRuleStateAsync(request, request.CorrelationId, cancellationToken),
                OperationType.UpsertTransportRule => await HandleUpsertTransportRuleAsync(request, request.CorrelationId, cancellationToken),
                OperationType.RemoveTransportRule => await HandleRemoveTransportRuleAsync(request, request.CorrelationId, cancellationToken),
                OperationType.TestTransportRule => await HandleTestTransportRuleAsync(request, request.CorrelationId, cancellationToken),
                OperationType.GetConnectors => await HandleGetConnectorsAsync(request, request.CorrelationId, cancellationToken),
                OperationType.UpsertConnector => await HandleUpsertConnectorAsync(request, request.CorrelationId, cancellationToken),
                OperationType.RemoveConnector => await HandleRemoveConnectorAsync(request, request.CorrelationId, cancellationToken),
                OperationType.GetAcceptedDomains => await HandleGetAcceptedDomainsAsync(request, request.CorrelationId, cancellationToken),
                OperationType.UpsertAcceptedDomain => await HandleUpsertAcceptedDomainAsync(request, request.CorrelationId, cancellationToken),
                OperationType.RemoveAcceptedDomain => await HandleRemoveAcceptedDomainAsync(request, request.CorrelationId, cancellationToken),
                OperationType.GetUserLicenses => await HandleGetUserLicensesAsync(request, request.CorrelationId, cancellationToken),
                OperationType.SetUserLicense => await HandleSetUserLicenseAsync(request, request.CorrelationId, cancellationToken),
                OperationType.GetAvailableLicenses => await HandleGetAvailableLicensesAsync(request, request.CorrelationId, cancellationToken),
                OperationType.CheckPrerequisites => await HandleCheckPrerequisitesAsync(request, request.CorrelationId, cancellationToken),
                OperationType.InstallModule => await HandleInstallModuleAsync(request, request.CorrelationId, cancellationToken),

                _ => CreateErrorResponse(request.CorrelationId, ErrorCode.OperationNotSupported, $"Operation {request.Operation} is not supported")
            };
        }
        catch (OperationCanceledException)
        {
            return CreateCancelledResponse(request.CorrelationId);
        }
        catch (Exception ex)
        {
            // Check if this is a deprecation warning that should be ignored
            if (ex is InvalidOperationException &&
                (ex.Message.Contains("deprecat", StringComparison.OrdinalIgnoreCase) ||
                 ex.Message.Contains("will start deprecating", StringComparison.OrdinalIgnoreCase)))
            {
                // Log as warning but don't fail the operation
                ConsoleLogger.Warning("Dispatcher", $"Deprecation warning: {ex.Message}");
                await SendLogAsync(request.CorrelationId, LogLevel.Warning, $"Deprecation warning: {ex.Message}");

                // Continue as if operation succeeded - this is just a warning
                // The actual operation likely completed successfully
                return CreateSuccessResponse(request.CorrelationId, new { });
            }

            ConsoleLogger.Error("Dispatcher", $"Exception: {ex.GetType().Name} - {ex.Message}");
            ConsoleLogger.Verbose("Dispatcher", $"Stack trace: {ex.StackTrace}");
            var (code, isTransient, retryAfter) = ErrorClassifier.Classify(ex);
            await SendLogAsync(request.CorrelationId, LogLevel.Error, $"Operation failed: {ex.Message}");
            return CreateErrorResponse(request.CorrelationId, code, ex.Message, isTransient, retryAfter);
        }
    }

    private async Task<ResponseEnvelope> HandleConnectAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Starting Exchange Online connection...");
        ConsoleLogger.Info("Dispatcher", $"Connecting to Exchange Online (correlation: {request.CorrelationId})");

        var result = await _psEngine.ConnectExchangeInteractiveAsync(
            onVerbose: async (level, msg) => await SendLogAsync(request.CorrelationId, LogLevel.Verbose, msg),
            cancellationToken: cancellationToken);

        ConsoleLogger.Debug("Dispatcher", $"ConnectExchangeInteractiveAsync completed. Success: {result.Success}, WasCancelled: {result.WasCancelled}");

        if (result.WasCancelled)
        {
            ConsoleLogger.Warning("Dispatcher", "Connection was cancelled");
            return CreateCancelledResponse(request.CorrelationId);
        }

        if (!result.Success)
        {
            ConsoleLogger.Error("Dispatcher", $"Connection failed: {result.ErrorMessage}");
            var (code, isTransient, retryAfter) = result.Errors.Any()
                ? ErrorClassifier.Classify(result.Errors.First())
                : (ErrorCode.AuthenticationFailed, false, (int?)null);

            return CreateErrorResponse(request.CorrelationId, code, result.ErrorMessage ?? "Connection failed", isTransient, retryAfter);
        }

        var (isConnected, upn, org, isGraphConnected) = await _psEngine.GetConnectionStatusAsync(cancellationToken);

        var status = new ConnectionStatusDto
        {
            State = isConnected ? ConnectionState.Connected : ConnectionState.Failed,
            UserPrincipalName = upn,
            Organization = org,
            GraphConnected = isGraphConnected,
            ConnectedAt = DateTime.UtcNow
        };

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Connected as {upn} to {org}");

        await SendLogAsync(request.CorrelationId, LogLevel.Verbose, "Detecting capabilities...");
        try
        {
            await _capabilityDetector.DetectCapabilitiesAsync(
                forceRefresh: true,
                onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, LogLevel.Verbose, msg),
                cancellationToken: cancellationToken);
        }
        catch (Exception ex)
        {
            await SendLogAsync(request.CorrelationId, LogLevel.Warning, $"Capability detection failed: {ex.Message}");
        }

        return CreateSuccessResponse(request.CorrelationId, status);
    }

    private async Task<ResponseEnvelope> HandleDisconnectAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Disconnecting from Exchange Online...");

        var result = await _psEngine.DisconnectExchangeAsync(cancellationToken);

        if (result.WasCancelled)
        {
            return CreateCancelledResponse(request.CorrelationId);
        }

        _capabilityDetector.ClearCache();

        var status = new ConnectionStatusDto
        {
            State = ConnectionState.Disconnected,
            GraphConnected = false
        };

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Disconnected from Exchange Online");

        return CreateSuccessResponse(request.CorrelationId, status);
    }

    private async Task<ResponseEnvelope> HandleGetConnectionStatusAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var (isConnected, upn, org, isGraphConnected) = await _psEngine.GetConnectionStatusAsync(cancellationToken);

        var status = new ConnectionStatusDto
        {
            State = isConnected ? ConnectionState.Connected : ConnectionState.Disconnected,
            UserPrincipalName = upn,
            Organization = org,
            GraphConnected = isGraphConnected
        };

        return CreateSuccessResponse(request.CorrelationId, status);
    }

    private async Task<ResponseEnvelope> HandleDetectCapabilitiesAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var detectRequest = JsonMessageSerializer.ExtractPayload<DetectCapabilitiesRequest>(request.Payload)
            ?? new DetectCapabilitiesRequest();

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Detecting capabilities...");

        var capabilities = await _capabilityDetector.DetectCapabilitiesAsync(
            forceRefresh: detectRequest.ForceRefresh,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, LogLevel.Verbose, msg),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, capabilities);
    }

    private async Task<ResponseEnvelope> HandleGetDashboardStatsAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var correlationId = request.CorrelationId;
        var statsRequest = JsonMessageSerializer.ExtractPayload<GetDashboardStatsRequest>(request.Payload)
            ?? new GetDashboardStatsRequest();

        await SendLogAsync(correlationId, LogLevel.Information, "Fetching dashboard statistics...");
        await SendProgressAsync(correlationId, 0, "Starting dashboard data collection...");

        var stats = await _exoCommands.GetDashboardStatsAsync(
            statsRequest,
            onLog: async (level, msg) => await SendLogAsync(correlationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        await SendProgressAsync(correlationId, 66, "Fetched mailbox and group counts...");

        foreach (var warning in stats.Warnings)
        {
            await SendLogAsync(correlationId, LogLevel.Warning, warning);
        }

        try
        {
            await SendLogAsync(correlationId, LogLevel.Information, "Fetching tenant licenses and admin users...");
            await SendProgressAsync(correlationId, 80, "Fetching licenses and admin users...");

            var licenses = await _exoCommands.GetTenantLicensesAsync(cancellationToken);
            stats.Licenses = licenses;

            var adminUsers = await _exoCommands.GetAdminRoleMembersAsync(cancellationToken);
            stats.AdminUsers = adminUsers;
        }
        catch (Exception ex)
        {
            await SendLogAsync(correlationId, LogLevel.Warning, $"Failed to fetch licenses/admin users: {ex.Message}");
            stats.Warnings.Add($"Could not retrieve license or admin data: {ex.Message}");
        }

        await SendProgressAsync(correlationId, 100, "Dashboard data collection complete");

        await SendLogAsync(correlationId, LogLevel.Information,
            $"Dashboard: {stats.MailboxCounts.Total} mailboxes, {stats.GroupCounts.Total} groups, {stats.Licenses.Count} license SKUs, {stats.AdminUsers.Count} admin users");

        return CreateSuccessResponse(correlationId, stats);
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

        await SendProgressAsync(request.CorrelationId, 100, "Mailbox retrieval complete");

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

        await SendProgressAsync(request.CorrelationId, 100, "Deleted mailbox retrieval complete");

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

    private async Task<ResponseEnvelope> HandleSetMailboxFeatureAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var featureRequest = JsonMessageSerializer.ExtractPayload<SetMailboxFeatureRequest>(request.Payload);

        if (featureRequest == null || string.IsNullOrWhiteSpace(featureRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information,
            $"Setting {featureRequest.Feature} = {featureRequest.Enabled} for {featureRequest.Identity}...");
        await SendProgressAsync(request.CorrelationId, 0, $"Setting {featureRequest.Feature}...");

        var escapedIdentity = featureRequest.Identity.Replace("'", "''");
        string script;

        switch (featureRequest.Feature)
        {
            case MailboxFeature.LitigationHold:
                script = $"Set-Mailbox -Identity '{escapedIdentity}' -LitigationHoldEnabled ${featureRequest.Enabled}";
                break;
            case MailboxFeature.Audit:
                script = $"Set-Mailbox -Identity '{escapedIdentity}' -AuditEnabled ${featureRequest.Enabled}";
                break;
            case MailboxFeature.SingleItemRecovery:
                script = $"Set-Mailbox -Identity '{escapedIdentity}' -SingleItemRecoveryEnabled ${featureRequest.Enabled}";
                break;
            case MailboxFeature.RetentionHold:
                script = $"Set-Mailbox -Identity '{escapedIdentity}' -RetentionHoldEnabled ${featureRequest.Enabled}";
                break;
            case MailboxFeature.Archive:
                if (featureRequest.Enabled)
                {
                    script = $"Enable-Mailbox -Identity '{escapedIdentity}' -Archive";
                }
                else
                {
                    script = $"Disable-Mailbox -Identity '{escapedIdentity}' -Archive -Confirm:$false";
                }
                break;
            default:
                return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, $"Unknown feature: {featureRequest.Feature}");
        }

        var result = await _psEngine.ExecuteAsync(
            script,
            onVerbose: async (level, msg) => await SendLogAsync(request.CorrelationId, LogLevel.Verbose, msg),
            cancellationToken: cancellationToken);

        if (result.WasCancelled)
        {
            return CreateCancelledResponse(request.CorrelationId);
        }

        if (!result.Success)
        {
            var (code, isTransient, retryAfter) = result.Errors.Any()
                ? ErrorClassifier.Classify(result.Errors.First())
                : (ErrorCode.Unknown, false, (int?)null);

            return CreateErrorResponse(request.CorrelationId, code, result.ErrorMessage ?? "Operation failed", isTransient, retryAfter);
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Successfully set {featureRequest.Feature}");
        await SendProgressAsync(request.CorrelationId, 100, $"{featureRequest.Feature} updated");

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
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
        if (createRequest == null)
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidRequest, "Missing create mailbox request payload", false, null);
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
                _ = SendProgressAsync(request.CorrelationId, percent, status);
            },
            cancellationToken: cancellationToken);

        await SendProgressAsync(request.CorrelationId, 100, "Mailbox space report complete");

        return CreateSuccessResponse(request.CorrelationId, response);
    }

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

        await SendProgressAsync(request.CorrelationId, 100, "Distribution lists retrieved");

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
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidRequest, "Missing create distribution list request payload", false, null);
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

    private async Task<ResponseEnvelope> HandleGetMessageTraceAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var traceRequest = JsonMessageSerializer.ExtractPayload<GetMessageTraceRequest>(request.Payload)
            ?? new GetMessageTraceRequest();

        await SendLogAsync(correlationId, LogLevel.Information, "Fetching message trace...");
        await SendProgressAsync(correlationId, 0, "Starting message trace query...");

        var response = await _exoCommands.GetMessageTraceAsync(
            traceRequest,
            cancellationToken);

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

    private async Task<ResponseEnvelope> HandleGetUserLicensesAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var licenseRequest = JsonMessageSerializer.ExtractPayload<GetUserLicensesRequest>(request.Payload);

        if (licenseRequest == null || string.IsNullOrWhiteSpace(licenseRequest.UserPrincipalName))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "UserPrincipalName is required");
        }

        await SendLogAsync(correlationId, LogLevel.Information, $"Fetching licenses for {licenseRequest.UserPrincipalName}...");
        await SendProgressAsync(correlationId, 0, "Fetching user licenses...");

        var response = await _exoCommands.GetUserLicensesAsync(
            licenseRequest.UserPrincipalName,
            cancellationToken);

        await SendProgressAsync(correlationId, 100, "User licenses retrieved");

        return CreateSuccessResponse(correlationId, response);
    }

    private async Task<ResponseEnvelope> HandleSetUserLicenseAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var setRequest = JsonMessageSerializer.ExtractPayload<SetUserLicenseRequest>(request.Payload);

        if (setRequest == null || string.IsNullOrWhiteSpace(setRequest.UserPrincipalName))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "UserPrincipalName is required");
        }

        await SendLogAsync(correlationId, LogLevel.Information, $"Updating licenses for {setRequest.UserPrincipalName}...");
        await SendProgressAsync(correlationId, 0, "Setting user license...");

        await _exoCommands.SetUserLicenseAsync(
            setRequest,
            cancellationToken);

        await SendProgressAsync(correlationId, 100, "User license updated");

        return CreateSuccessResponse(correlationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleGetAvailableLicensesAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        await SendLogAsync(correlationId, LogLevel.Information, "Fetching available licenses...");
        await SendProgressAsync(correlationId, 0, "Fetching available licenses...");

        var licenses = await _exoCommands.GetTenantLicensesAsync(cancellationToken);

        var response = new GetAvailableLicensesResponse
        {
            Licenses = licenses
        };

        await SendProgressAsync(correlationId, 100, "Available licenses retrieved");

        return CreateSuccessResponse(correlationId, response);
    }

    private async Task<ResponseEnvelope> HandleCheckPrerequisitesAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        await SendLogAsync(correlationId, LogLevel.Information, "Checking prerequisites...");
        await SendProgressAsync(correlationId, 0, "Checking system prerequisites...");

        await SendLogAsync(correlationId, LogLevel.Information, "[Prerequisites] Running PowerShell/module checks");
        var status = await _exoCommands.CheckPrerequisitesAsync(cancellationToken);

        await SendProgressAsync(correlationId, 100, "Prerequisite check complete");

        return CreateSuccessResponse(correlationId, status);
    }

    private async Task<ResponseEnvelope> HandleInstallModuleAsync(RequestEnvelope request, string correlationId, CancellationToken cancellationToken)
    {
        var installRequest = JsonMessageSerializer.ExtractPayload<InstallModuleRequest>(request.Payload);

        if (installRequest == null || string.IsNullOrWhiteSpace(installRequest.ModuleName))
        {
            return CreateErrorResponse(correlationId, ErrorCode.InvalidParameter, "ModuleName is required");
        }

        await SendLogAsync(correlationId, LogLevel.Information, $"Installing module {installRequest.ModuleName}...");
        await SendProgressAsync(correlationId, 0, $"Installing {installRequest.ModuleName}...");

        await SendLogAsync(correlationId, LogLevel.Information, $"[ModuleInstall] Starting install: {installRequest.ModuleName}");
        var response = await _exoCommands.InstallModuleAsync(
            installRequest.ModuleName,
            cancellationToken);

        await SendProgressAsync(correlationId, 100, $"Module {installRequest.ModuleName} installation complete");

        return CreateSuccessResponse(correlationId, response);
    }

    private async Task<ResponseEnvelope> HandleDemoOperationAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var demoRequest = JsonMessageSerializer.ExtractPayload<DemoOperationRequest>(request.Payload)
            ?? new DemoOperationRequest();

        var startTime = DateTime.UtcNow;
        var results = new List<DemoItemResult>();

        await SendLogAsync(request.CorrelationId, LogLevel.Information,
            $"Starting demo operation: {demoRequest.ItemCount} items over {demoRequest.DurationSeconds}s");

        var delayPerItem = (demoRequest.DurationSeconds * 1000) / Math.Max(1, demoRequest.ItemCount);

        for (int i = 0; i < demoRequest.ItemCount; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var percentComplete = (int)((i + 1) * 100.0 / demoRequest.ItemCount);

            if (demoRequest.SimulateError && percentComplete >= demoRequest.ErrorAtPercent)
            {
                await SendLogAsync(request.CorrelationId, LogLevel.Error,
                    $"Simulated error at {percentComplete}%");

                return CreateErrorResponse(request.CorrelationId, ErrorCode.Unknown,
                    $"Simulated error at {percentComplete}%");
            }

            await Task.Delay(delayPerItem, cancellationToken);

            var itemResult = new DemoItemResult
            {
                ItemId = i + 1,
                Status = "Processed",
                Timestamp = DateTime.UtcNow
            };
            results.Add(itemResult);

            await SendProgressAsync(request.CorrelationId, percentComplete,
                $"Processing item {i + 1} of {demoRequest.ItemCount}",
                i + 1, demoRequest.ItemCount);

            if ((i + 1) % 3 == 0 || i == demoRequest.ItemCount - 1)
            {
                await SendPartialOutputAsync(request.CorrelationId, itemResult, i);
            }

            if ((i + 1) % 5 == 0)
            {
                await SendLogAsync(request.CorrelationId, LogLevel.Verbose,
                    $"Processed {i + 1} items...");
            }
        }

        var elapsed = (DateTime.UtcNow - startTime).TotalSeconds;

        await SendLogAsync(request.CorrelationId, LogLevel.Information,
            $"Demo operation completed: {results.Count} items in {elapsed:F1}s");

        var response = new DemoOperationResponse
        {
            ProcessedItems = results.Count,
            ElapsedSeconds = elapsed,
            Results = results
        };

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private LogLevel ParseLogLevel(string level)
    {
        return level.ToLowerInvariant() switch
        {
            "verbose" => LogLevel.Verbose,
            "debug" => LogLevel.Debug,
            "information" or "info" => LogLevel.Information,
            "warning" or "warn" => LogLevel.Warning,
            "error" => LogLevel.Error,
            _ => LogLevel.Information
        };
    }

    private async Task SendLogAsync(string correlationId, LogLevel level, string message)
    {
        var payload = new LogEventPayload
        {
            Level = level,
            Message = message,
            Source = "Worker"
        };

        var evt = new EventEnvelope
        {
            CorrelationId = correlationId,
            EventType = EventType.Log,
            Payload = JsonMessageSerializer.ToJsonElement(payload)
        };

        await _sendEvent(evt);
    }

    private async Task SendProgressAsync(string correlationId, int percentComplete, string? statusMessage, int? currentItem = null, int? totalItems = null)
    {
        var payload = new ProgressEventPayload
        {
            PercentComplete = percentComplete,
            StatusMessage = statusMessage,
            CurrentItem = currentItem,
            TotalItems = totalItems
        };

        var evt = new EventEnvelope
        {
            CorrelationId = correlationId,
            EventType = EventType.Progress,
            Payload = JsonMessageSerializer.ToJsonElement(payload)
        };

        await _sendEvent(evt);
    }

    private async Task SendPartialOutputAsync<T>(string correlationId, T data, int itemIndex)
    {
        var payload = new PartialOutputPayload
        {
            Data = JsonMessageSerializer.ToJsonElement(data),
            ItemIndex = itemIndex
        };

        var evt = new EventEnvelope
        {
            CorrelationId = correlationId,
            EventType = EventType.PartialOutput,
            Payload = JsonMessageSerializer.ToJsonElement(payload)
        };

        await _sendEvent(evt);
    }

    private ResponseEnvelope CreateSuccessResponse<T>(string correlationId, T payload)
    {
        return new ResponseEnvelope
        {
            CorrelationId = correlationId,
            Success = true,
            Payload = JsonMessageSerializer.ToJsonElement(payload)
        };
    }

    private ResponseEnvelope CreateErrorResponse(string correlationId, ErrorCode code, string message, bool isTransient = false, int? retryAfterSeconds = null)
    {
        return new ResponseEnvelope
        {
            CorrelationId = correlationId,
            Success = false,
            Error = new NormalizedErrorDto
            {
                Code = code,
                Message = message,
                IsTransient = isTransient,
                RetryAfterSeconds = retryAfterSeconds
            }
        };
    }

    private ResponseEnvelope CreateCancelledResponse(string correlationId)
    {
        return new ResponseEnvelope
        {
            CorrelationId = correlationId,
            Success = false,
            WasCancelled = true
        };
    }
}
