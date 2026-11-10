using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Worker.PowerShell;

namespace OnlyExo365.Worker.Operations;

public partial class OperationDispatcher
{
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

        stats.CorrelationId = correlationId;
        await SendProgressAsync(correlationId, 85, "Dashboard data collected...");

        foreach (var warning in stats.Warnings)
        {
            await SendLogAsync(correlationId, LogLevel.Warning, warning);
        }

        await SendProgressAsync(correlationId, 100, "Dashboard data collection complete");

        await SendLogAsync(correlationId, LogLevel.Information,
            $"Dashboard: {stats.MailboxCounts.Total} mailboxes, {stats.GroupCounts.Total} groups, {stats.Licenses.Count} license SKUs, {stats.AdminUsers.Count} admin users");

        return CreateSuccessResponse(correlationId, stats);
    }

    private async Task<ResponseEnvelope> HandleGetContactsAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var contactsRequest = JsonMessageSerializer.ExtractPayload<GetContactsRequest>(request.Payload)
            ?? new GetContactsRequest();

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Fetching contacts...");
        await SendProgressAsync(request.CorrelationId, 0, "Starting contacts retrieval...");

        var response = await _exoCommands.GetContactsAsync(
            contactsRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        await SendProgressAsync(
            request.CorrelationId,
            100,
            FormatListProgressStatus(response.Contacts.Count, response.TotalCount, "Analizzati"),
            response.Contacts.Count,
            response.TotalCount);

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleGetContactDetailsAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var detailsRequest = JsonMessageSerializer.ExtractPayload<GetContactDetailsRequest>(request.Payload);

        if (detailsRequest == null || string.IsNullOrWhiteSpace(detailsRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Fetching contact details for {detailsRequest.Identity}...");

        var details = await _exoCommands.GetContactDetailsAsync(
            detailsRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, details);
    }

    private async Task<ResponseEnvelope> HandleUpsertContactAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var upsertRequest = JsonMessageSerializer.ExtractPayload<UpsertContactRequest>(request.Payload);
        if (upsertRequest != null)
        {
            WorkerSecretResolver.Resolve(upsertRequest);
        }

        if (upsertRequest == null ||
            string.IsNullOrWhiteSpace(upsertRequest.DisplayName) ||
            string.IsNullOrWhiteSpace(upsertRequest.Alias) ||
            string.IsNullOrWhiteSpace(upsertRequest.PrimarySmtpAddress) ||
            string.IsNullOrWhiteSpace(upsertRequest.ExternalEmailAddress))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "DisplayName, Alias, PrimarySmtpAddress and ExternalEmailAddress are required");
        }

        if (string.Equals(upsertRequest.ContactKind, "MailUser", StringComparison.OrdinalIgnoreCase) &&
            (string.IsNullOrWhiteSpace(upsertRequest.Identity) &&
             (string.IsNullOrWhiteSpace(upsertRequest.UserPrincipalName) || string.IsNullOrWhiteSpace(upsertRequest.Password))))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "UserPrincipalName and Password are required when creating a MailUser");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Saving contact {upsertRequest.DisplayName} ({upsertRequest.ContactKind})...");

        await _exoCommands.UpsertContactAsync(upsertRequest, cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleRemoveContactAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var removeRequest = JsonMessageSerializer.ExtractPayload<RemoveContactRequest>(request.Payload);

        if (removeRequest == null || string.IsNullOrWhiteSpace(removeRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Warning, $"Removing contact: {removeRequest.Identity} ({removeRequest.ContactKind})");

        await _exoCommands.RemoveContactAsync(removeRequest, cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleGetResourceMailboxesAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var resourceRequest = JsonMessageSerializer.ExtractPayload<GetResourceMailboxesRequest>(request.Payload)
            ?? new GetResourceMailboxesRequest();

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Fetching resource mailboxes...");
        await SendProgressAsync(request.CorrelationId, 0, "Starting resources retrieval...");

        var response = await _exoCommands.GetResourceMailboxesAsync(
            resourceRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        await SendProgressAsync(
            request.CorrelationId,
            100,
            FormatListProgressStatus(response.Resources.Count, response.TotalCount, "Analizzati"),
            response.Resources.Count,
            response.TotalCount);

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleGetResourceMailboxDetailsAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var detailsRequest = JsonMessageSerializer.ExtractPayload<GetResourceMailboxDetailsRequest>(request.Payload);

        if (detailsRequest == null || string.IsNullOrWhiteSpace(detailsRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Fetching resource details for {detailsRequest.Identity}...");

        var details = await _exoCommands.GetResourceMailboxDetailsAsync(
            detailsRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, details);
    }

    private async Task<ResponseEnvelope> HandleUpsertResourceMailboxAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var upsertRequest = JsonMessageSerializer.ExtractPayload<UpsertResourceMailboxRequest>(request.Payload);

        if (upsertRequest == null ||
            string.IsNullOrWhiteSpace(upsertRequest.DisplayName) ||
            string.IsNullOrWhiteSpace(upsertRequest.Alias) ||
            string.IsNullOrWhiteSpace(upsertRequest.PrimarySmtpAddress))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "DisplayName, Alias and PrimarySmtpAddress are required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Saving resource mailbox {upsertRequest.DisplayName}...");

        var response = await _exoCommands.UpsertResourceMailboxAsync(
            upsertRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleGetPublicFoldersAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var publicFoldersRequest = JsonMessageSerializer.ExtractPayload<GetPublicFoldersRequest>(request.Payload)
            ?? new GetPublicFoldersRequest();

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Fetching public folders...");
        await SendProgressAsync(request.CorrelationId, 0, "Starting public folders retrieval...");

        var response = await _exoCommands.GetPublicFoldersAsync(
            publicFoldersRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        await SendProgressAsync(
            request.CorrelationId,
            100,
            FormatListProgressStatus(response.Folders.Count, response.TotalCount, "Analizzate"),
            response.Folders.Count,
            response.TotalCount);

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleGetPublicFolderDetailsAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var detailsRequest = JsonMessageSerializer.ExtractPayload<GetPublicFolderDetailsRequest>(request.Payload);

        if (detailsRequest == null || string.IsNullOrWhiteSpace(detailsRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Fetching public folder details for {detailsRequest.Identity}...");

        var details = await _exoCommands.GetPublicFolderDetailsAsync(
            detailsRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, details);
    }

    private async Task<ResponseEnvelope> HandleUpsertPublicFolderAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var upsertRequest = JsonMessageSerializer.ExtractPayload<UpsertPublicFolderRequest>(request.Payload);

        if (upsertRequest == null || string.IsNullOrWhiteSpace(upsertRequest.Name))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Name is required");
        }

        if (upsertRequest.MailEnabled && string.IsNullOrWhiteSpace(upsertRequest.Alias))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Alias is required when MailEnabled is true");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Saving public folder {upsertRequest.Name}...");

        var response = await _exoCommands.UpsertPublicFolderAsync(
            upsertRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleGetMobileDevicesAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var devicesRequest = JsonMessageSerializer.ExtractPayload<GetMobileDevicesRequest>(request.Payload)
            ?? new GetMobileDevicesRequest();

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Fetching mobile devices...");
        await SendProgressAsync(request.CorrelationId, 0, "Starting mobile device retrieval...");

        var response = await _exoCommands.GetMobileDevicesAsync(
            devicesRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        await SendProgressAsync(
            request.CorrelationId,
            100,
            FormatListProgressStatus(response.Devices.Count, response.TotalCount, "Analizzati"),
            response.Devices.Count,
            response.TotalCount);

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleSetPublicFolderClientPermissionAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var permissionRequest = JsonMessageSerializer.ExtractPayload<SetPublicFolderClientPermissionRequest>(request.Payload);

        if (permissionRequest == null ||
            string.IsNullOrWhiteSpace(permissionRequest.Identity) ||
            string.IsNullOrWhiteSpace(permissionRequest.User))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity and User are required");
        }

        if (permissionRequest.Action != PermissionAction.Remove && permissionRequest.AccessRights.Count == 0)
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "AccessRights are required for add or modify actions");
        }

        await SendLogAsync(
            request.CorrelationId,
            LogLevel.Information,
            $"{permissionRequest.Action} public folder client permission for {permissionRequest.User} on {permissionRequest.Identity}...");

        await _exoCommands.SetPublicFolderClientPermissionAsync(
            permissionRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleRemovePublicFolderAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var removeRequest = JsonMessageSerializer.ExtractPayload<RemovePublicFolderRequest>(request.Payload);

        if (removeRequest == null || string.IsNullOrWhiteSpace(removeRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Warning, $"Removing public folder {removeRequest.Identity}...");

        await _exoCommands.RemovePublicFolderAsync(
            removeRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleGetMobileDeviceDetailsAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var detailsRequest = JsonMessageSerializer.ExtractPayload<GetMobileDeviceDetailsRequest>(request.Payload);

        if (detailsRequest == null || string.IsNullOrWhiteSpace(detailsRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Fetching mobile device details for {detailsRequest.Identity}...");
        await SendProgressAsync(request.CorrelationId, 0, "Starting mobile device details retrieval...");

        var response = await _exoCommands.GetMobileDeviceDetailsAsync(
            detailsRequest,
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            onProgress: async (current, total, status) =>
            {
                var safeTotal = total <= 0 ? 1 : total;
                var percent = (int)Math.Round(current * 100d / safeTotal, MidpointRounding.AwayFromZero);
                await SendProgressAsync(request.CorrelationId, percent, status, current, total);
            },
            cancellationToken: cancellationToken);

        await SendProgressAsync(request.CorrelationId, 100, "Mobile device details retrieval complete");

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleGetMobileDeviceMailboxPoliciesAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Fetching mobile device mailbox policies...");

        var response = await _exoCommands.GetMobileDeviceMailboxPoliciesAsync(
            onLog: async (level, msg) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), msg),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleSetMobileDeviceAccessStateAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var accessRequest = JsonMessageSerializer.ExtractPayload<SetMobileDeviceAccessStateRequest>(request.Payload);

        if (accessRequest == null ||
            string.IsNullOrWhiteSpace(accessRequest.MailboxIdentity) ||
            string.IsNullOrWhiteSpace(accessRequest.DeviceId) ||
            string.IsNullOrWhiteSpace(accessRequest.AccessState))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "MailboxIdentity, DeviceId and AccessState are required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Setting mobile device access state: {accessRequest.DeviceId} => {accessRequest.AccessState}");
        await _exoCommands.SetMobileDeviceAccessStateAsync(accessRequest, cancellationToken);
        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleClearMobileDeviceAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var clearRequest = JsonMessageSerializer.ExtractPayload<ClearMobileDeviceRequest>(request.Payload);

        if (clearRequest == null || string.IsNullOrWhiteSpace(clearRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Warning, $"Issuing remote wipe for mobile device: {clearRequest.Identity}");
        await _exoCommands.ClearMobileDeviceAsync(clearRequest, cancellationToken);
        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleSetMobileDeviceMailboxPolicyAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var policyRequest = JsonMessageSerializer.ExtractPayload<SetMobileDeviceMailboxPolicyRequest>(request.Payload);

        if (policyRequest == null || string.IsNullOrWhiteSpace(policyRequest.MailboxIdentity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "MailboxIdentity is required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Setting mobile device mailbox policy for {policyRequest.MailboxIdentity}...");
        await _exoCommands.SetMobileDeviceMailboxPolicyAsync(policyRequest, cancellationToken);
        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleGetMigrationBatchesAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var batchesRequest = JsonMessageSerializer.ExtractPayload<GetMigrationBatchesRequest>(request.Payload)
            ?? new GetMigrationBatchesRequest();

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Fetching migration batches...");
        await SendProgressAsync(request.CorrelationId, 0, "Starting migration batch retrieval...");

        var response = await _exoCommands.GetMigrationBatchesAsync(
            batchesRequest,
            onLog: async (level, message) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), message),
            cancellationToken: cancellationToken);

        await SendProgressAsync(
            request.CorrelationId,
            100,
            FormatListProgressStatus(response.Batches.Count, response.TotalCount, "Analizzati"),
            response.Batches.Count,
            response.TotalCount);

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleGetMigrationEndpointsAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var endpointsRequest = JsonMessageSerializer.ExtractPayload<GetMigrationEndpointsRequest>(request.Payload)
            ?? new GetMigrationEndpointsRequest();

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Fetching migration endpoints...");

        var response = await _exoCommands.GetMigrationEndpointsAsync(
            endpointsRequest,
            onLog: async (level, message) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), message),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleGetMigrationBatchDetailsAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var detailsRequest = JsonMessageSerializer.ExtractPayload<GetMigrationBatchDetailsRequest>(request.Payload);

        if (detailsRequest == null || string.IsNullOrWhiteSpace(detailsRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Fetching migration batch details for {detailsRequest.Identity}...");

        var response = await _exoCommands.GetMigrationBatchDetailsAsync(
            detailsRequest,
            onLog: async (level, message) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), message),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleUpsertMigrationEndpointAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var upsertRequest = JsonMessageSerializer.ExtractPayload<UpsertMigrationEndpointRequest>(request.Payload);
        if (upsertRequest != null)
        {
            WorkerSecretResolver.Resolve(upsertRequest);
        }

        if (upsertRequest == null || string.IsNullOrWhiteSpace(upsertRequest.Name))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Name is required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Saving migration endpoint {upsertRequest.Name}...");
        await _exoCommands.UpsertMigrationEndpointAsync(upsertRequest, cancellationToken);
        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleTestMigrationEndpointAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var testRequest = JsonMessageSerializer.ExtractPayload<TestMigrationEndpointRequest>(request.Payload);
        if (testRequest != null)
        {
            WorkerSecretResolver.Resolve(testRequest);
        }

        if (testRequest == null)
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Migration endpoint test request is required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Testing migration endpoint...");
        var response = await _exoCommands.TestMigrationEndpointAsync(testRequest, cancellationToken);
        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleGetMigrationBatchPreflightAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var preflightRequest = JsonMessageSerializer.ExtractPayload<GetMigrationBatchPreflightRequest>(request.Payload);
        if (preflightRequest == null ||
            string.IsNullOrWhiteSpace(preflightRequest.Name) ||
            string.IsNullOrWhiteSpace(preflightRequest.EndpointIdentity) ||
            string.IsNullOrWhiteSpace(preflightRequest.CsvFilePath))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Name, EndpointIdentity and CsvFilePath are required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Running migration batch preflight for {preflightRequest.Name}...");
        var response = await _exoCommands.GetMigrationBatchPreflightAsync(preflightRequest, cancellationToken);
        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleCreateMigrationBatchAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var createRequest = JsonMessageSerializer.ExtractPayload<CreateMigrationBatchRequest>(request.Payload);
        if (createRequest == null ||
            string.IsNullOrWhiteSpace(createRequest.Name) ||
            string.IsNullOrWhiteSpace(createRequest.EndpointIdentity) ||
            string.IsNullOrWhiteSpace(createRequest.CsvFilePath))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Name, EndpointIdentity and CsvFilePath are required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Creating migration batch {createRequest.Name}...");
        await _exoCommands.CreateMigrationBatchAsync(createRequest, cancellationToken);
        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleStartMigrationBatchAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var startRequest = JsonMessageSerializer.ExtractPayload<StartMigrationBatchRequest>(request.Payload);

        if (startRequest == null || string.IsNullOrWhiteSpace(startRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Starting migration batch {startRequest.Identity}...");
        await _exoCommands.StartMigrationBatchAsync(startRequest, cancellationToken);
        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleCompleteMigrationBatchAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var completeRequest = JsonMessageSerializer.ExtractPayload<CompleteMigrationBatchRequest>(request.Payload);

        if (completeRequest == null || string.IsNullOrWhiteSpace(completeRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Completing migration batch {completeRequest.Identity}...");
        await _exoCommands.CompleteMigrationBatchAsync(completeRequest, cancellationToken);
        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleRemoveMigrationBatchAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var removeRequest = JsonMessageSerializer.ExtractPayload<RemoveMigrationBatchRequest>(request.Payload);

        if (removeRequest == null || string.IsNullOrWhiteSpace(removeRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Warning, $"Removing migration batch {removeRequest.Identity}...");
        await _exoCommands.RemoveMigrationBatchAsync(removeRequest, cancellationToken);
        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleGetRoleGroupsAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var roleGroupsRequest = JsonMessageSerializer.ExtractPayload<GetRoleGroupsRequest>(request.Payload)
            ?? new GetRoleGroupsRequest();

        await SendLogAsync(request.CorrelationId, LogLevel.Information, "Fetching role groups...");
        await SendProgressAsync(request.CorrelationId, 0, "Starting role group retrieval...");

        var response = await _exoPermissionCommands.GetRoleGroupsAsync(
            roleGroupsRequest,
            onLog: async (level, message) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), message),
            cancellationToken: cancellationToken);

        await SendProgressAsync(
            request.CorrelationId,
            100,
            FormatListProgressStatus(response.RoleGroups.Count, response.TotalCount, "Analizzati"),
            response.RoleGroups.Count,
            response.TotalCount);
        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleGetRoleGroupDetailsAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var detailsRequest = JsonMessageSerializer.ExtractPayload<GetRoleGroupDetailsRequest>(request.Payload);
        if (detailsRequest == null || string.IsNullOrWhiteSpace(detailsRequest.Identity))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity is required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Fetching role group details for {detailsRequest.Identity}...");

        var response = await _exoPermissionCommands.GetRoleGroupDetailsAsync(
            detailsRequest,
            onLog: async (level, message) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), message),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, response);
    }

    private async Task<ResponseEnvelope> HandleUpsertRoleGroupAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var upsertRequest = JsonMessageSerializer.ExtractPayload<UpsertRoleGroupRequest>(request.Payload);
        if (upsertRequest == null || string.IsNullOrWhiteSpace(upsertRequest.Name))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Name is required");
        }

        await SendLogAsync(request.CorrelationId, LogLevel.Information, $"Saving role group {upsertRequest.Name}...");
        await _exoPermissionCommands.UpsertRoleGroupAsync(
            upsertRequest,
            onLog: async (level, message) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), message),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }

    private async Task<ResponseEnvelope> HandleModifyRoleGroupMemberAsync(RequestEnvelope request, CancellationToken cancellationToken)
    {
        var modifyRequest = JsonMessageSerializer.ExtractPayload<ModifyRoleGroupMemberRequest>(request.Payload);
        if (modifyRequest == null || string.IsNullOrWhiteSpace(modifyRequest.Identity) || string.IsNullOrWhiteSpace(modifyRequest.Member))
        {
            return CreateErrorResponse(request.CorrelationId, ErrorCode.InvalidParameter, "Identity and member are required");
        }

        await SendLogAsync(
            request.CorrelationId,
            LogLevel.Information,
            $"{modifyRequest.Action} role group member {modifyRequest.Member} on {modifyRequest.Identity}...");

        await _exoPermissionCommands.ModifyRoleGroupMemberAsync(
            modifyRequest,
            onLog: async (level, message) => await SendLogAsync(request.CorrelationId, ParseLogLevel(level), message),
            cancellationToken: cancellationToken);

        return CreateSuccessResponse(request.CorrelationId, new { Success = true });
    }
}

