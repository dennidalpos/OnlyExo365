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
        Console.WriteLine($"[OperationDispatcher] Dispatching operation: {request.Operation}");
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
                OperationType.GetMailboxSpaceReport => await HandleGetMailboxSpaceReportAsync(request, cancellationToken),

                                     
                OperationType.GetDistributionLists => await HandleGetDistributionListsAsync(request, cancellationToken),
                OperationType.GetDistributionListDetails => await HandleGetDistributionListDetailsAsync(request, cancellationToken),
                OperationType.GetGroupMembers => await HandleGetGroupMembersAsync(request, cancellationToken),
                OperationType.ModifyGroupMember => await HandleModifyGroupMemberAsync(request, cancellationToken),
                OperationType.PreviewDynamicGroupMembers => await HandlePreviewDynamicGroupMembersAsync(request, cancellationToken),
                OperationType.SetDistributionListSettings => await HandleSetDistributionListSettingsAsync(request, cancellationToken),

                _ => CreateErrorResponse(request.CorrelationId, ErrorCode.OperationNotSupported, $"Operation {request.Operation} is not supported")
            };
        }
        catch (OperationCanceledException)
        {
            return CreateCancelledResponse(request.CorrelationId);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[OperationDispatcher] Exception in DispatchAsync: {ex.GetType().Name} - {ex.Message}");
            Console.WriteLine($"[OperationDispatcher] Stack trace: {ex.StackTrace}");
            var (code, isTransient, retryAfter) = ErrorClassifier.Classify(ex);
            await SendLogAsync(request.CorrelationId, LogLevel.Error, $"Operation failed: {ex.Message}");
            return CreateErrorResponse(request.CorrelationId, code, ex.Message, isTransient, retryAfter);
        }
    }

    #region Connection Handlers

    private async Task<ResponseEnvelope> HandleConnectAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Starting Exchange Online connection...");
        Console.WriteLine($"[OperationDispatcher] Calling ConnectExchangeInteractiveAsync for correlation {request.CorrelationId}");

        var result = await _psEngine.ConnectExchangeInteractiveAsync(
            onVerbose: async (level, msg) => await SendLogAsync(request.CorrelationId, LogLevel.Verbose, msg),
            cancellationToken: cancellationToken);

        Console.WriteLine($"[OperationDispatcher] ConnectExchangeInteractiveAsync completed. Success: {result.Success}, WasCancelled: {result.WasCancelled}");

        if (result.WasCancelled)
        {
            Console.WriteLine($"[OperationDispatcher] Connection was cancelled");
            return CreateCancelledResponse(request.CorrelationId);
        }

        if (!result.Success)
        {
            Console.WriteLine($"[OperationDispatcher] Connection failed: {result.ErrorMessage}");
            var (code, isTransient, retryAfter) = result.Errors.Any()
                ? ErrorClassifier.Classify(result.Errors.First())
                : (ErrorCode.AuthenticationFailed, false, (int?)null);

            return CreateErrorResponse(request.CorrelationId, code, result.ErrorMessage ?? "Connection failed", isTransient, retryAfter);
        }

                              
        var (isConnected, upn, org) = await _psEngine.GetConnectionStatusAsync(cancellationToken);

        var status = new ConnectionStatusDto
        {
            State = isConnected ? ConnectionState.Connected : ConnectionState.Failed,
            UserPrincipalName = upn,
            Organization = org,
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
            State = ConnectionState.Disconnected
        };

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Disconnected from Exchange Online");

        return CreateSuccessResponse(request.CorrelationId, status);
    }

    private async Task<ResponseEnvelope> HandleGetConnectionStatusAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var (isConnected, upn, org) = await _psEngine.GetConnectionStatusAsync(cancellationToken);

        var status = new ConnectionStatusDto
        {
            State = isConnected ? ConnectionState.Connected : ConnectionState.Disconnected,
            UserPrincipalName = upn,
            Organization = org
        };

        return CreateSuccessResponse(request.CorrelationId, status);
    }

    #endregion

    #region Capability Detection

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

    #endregion

    #region Dashboard

    private async Task<ResponseEnvelope> HandleGetDashboardStatsAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var statsRequest = JsonMessageSerializer.ExtractPayload<GetDashboardStatsRequest>(request.Payload)
            ?? new GetDashboardStatsRequest();

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Fetching dashboard statistics...");

        var stats = await _exoCommands.GetDashboardStatsAsync(
            statsRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

                          
        foreach (var warning in stats.Warnings)
        {
            await SendLogAsync(request.CorrelationId, LogLevel.Warning, warning);
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information,
            $"Dashboard: {stats.MailboxCounts.Total} mailboxes, {stats.GroupCounts.Total} groups");

        return CreateSuccessResponse(request.CorrelationId, stats);
    }

    #endregion

    #region Mailbox Handlers

    private async Task<ResponseEnvelope> HandleGetMailboxesAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var mailboxRequest = JsonMessageSerializer.ExtractPayload<GetMailboxesRequest>(request.Payload)
            ?? new GetMailboxesRequest();

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Fetching mailboxes...");

        var response = await _exoCommands.GetMailboxesAsync(
            mailboxRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            onPartialOutput: async (item) => await SendPartialOutputAsync(request.CorrelationId, item, 0),
            cancellationToken: cancellationToken);

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

        var details = await _exoCommands.GetMailboxDetailsAsync(
            detailsRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, details);
    }

    private async Task<ResponseEnvelope> HandleGetRetentionPoliciesAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        _ = JsonMessageSerializer.ExtractPayload<GetRetentionPoliciesRequest>(request.Payload)
            ?? new GetRetentionPoliciesRequest();

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Fetching retention policies...");

        var policies = await _exoCommands.GetRetentionPoliciesAsync(
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        var response = new GetRetentionPoliciesResponse
        {
            Policies = policies
        };

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleSetRetentionPolicyAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var setRequest = JsonMessageSerializer.ExtractPayload<SetRetentionPolicyRequest>(request.Payload);

        if (setRequest == null || string.IsNullOrWhiteSpace(setRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await _exoCommands.SetRetentionPolicyAsync(
            setRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

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

        var permissions = await _exoCommands.GetMailboxPermissionsAsync(
            permRequest.Identity,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

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

        await _exoCommands.SetMailboxPermissionAsync(
            setRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

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

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleUpdateMailboxSettingsAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var settingsRequest = JsonMessageSerializer.ExtractPayload<UpdateMailboxSettingsRequest>(request.Payload);

        if (settingsRequest == null || string.IsNullOrWhiteSpace(settingsRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await _exoCommands.SetMailboxSettingsAsync(
            settingsRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleSetMailboxAutoReplyConfigurationAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var autoReplyRequest = JsonMessageSerializer.ExtractPayload<SetMailboxAutoReplyConfigurationRequest>(request.Payload);

        if (autoReplyRequest == null || string.IsNullOrWhiteSpace(autoReplyRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await _exoCommands.SetMailboxAutoReplyConfigurationAsync(
            autoReplyRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleConvertMailboxToSharedAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var convertRequest = JsonMessageSerializer.ExtractPayload<ConvertMailboxToSharedRequest>(request.Payload);

        if (convertRequest == null || string.IsNullOrWhiteSpace(convertRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await _exoCommands.ConvertMailboxToSharedAsync(
            convertRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleConvertMailboxToRegularAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var convertRequest = JsonMessageSerializer.ExtractPayload<ConvertMailboxToRegularRequest>(request.Payload);

        if (convertRequest == null || string.IsNullOrWhiteSpace(convertRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await _exoCommands.ConvertMailboxToRegularAsync(
            convertRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleGetMailboxSpaceReportAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var reportRequest = JsonMessageSerializer.ExtractPayload<GetMailboxSpaceReportRequest>(request.Payload)
            ?? new GetMailboxSpaceReportRequest();

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Fetching mailbox space report...");

        var response = await _exoCommands.GetMailboxSpaceReportAsync(
            reportRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    #endregion

    #region Distribution List Handlers

    private async Task<ResponseEnvelope> HandleGetDistributionListsAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var listRequest = JsonMessageSerializer.ExtractPayload<GetDistributionListsRequest>(request.Payload)
            ?? new GetDistributionListsRequest();

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Fetching distribution lists...");

        var response = await _exoGroupCommands.GetDistributionListsAsync(
            listRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            onPartialOutput: async (item) => await SendPartialOutputAsync(request.CorrelationId, item, 0),
            cancellationToken: cancellationToken);

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

        var details = await _exoGroupCommands.GetDistributionListDetailsAsync(
            detailsRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

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

        var members = await _exoGroupCommands.GetGroupMembersPageAsync(
            membersRequest.Identity,
            membersRequest.GroupType,
            membersRequest.Skip,
            membersRequest.PageSize,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, members);
    }

    private async Task<ResponseEnvelope> HandleModifyGroupMemberAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var modifyRequest = JsonMessageSerializer.ExtractPayload<ModifyGroupMemberRequest>(request.Payload);

        if (modifyRequest == null || string.IsNullOrWhiteSpace(modifyRequest.Identity) || string.IsNullOrWhiteSpace(modifyRequest.Member))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity and Member are required");
        }

        await _exoGroupCommands.ModifyGroupMemberAsync(
            modifyRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

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

        var response = await _exoGroupCommands.PreviewDynamicGroupMembersAsync(
            previewRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleSetDistributionListSettingsAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var settingsRequest = JsonMessageSerializer.ExtractPayload<SetDistributionListSettingsRequest>(request.Payload);

        if (settingsRequest == null || string.IsNullOrWhiteSpace(settingsRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await _exoGroupCommands.SetDistributionListSettingsAsync(
            settingsRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    #endregion

    #region Demo Handler

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

    #endregion

    #region Event Helpers

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

    #endregion

    #region Response Helpers

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

    #endregion
}
