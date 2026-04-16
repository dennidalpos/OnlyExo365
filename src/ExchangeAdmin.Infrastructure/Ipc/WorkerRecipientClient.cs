using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;

namespace ExchangeAdmin.Infrastructure.Ipc;

internal sealed class WorkerRecipientClient
{
    private readonly WorkerClientRuntime _runtime;

    public WorkerRecipientClient(WorkerClientRuntime runtime)
    {
        _runtime = runtime;
    }

    public Task<Result<GetContactsResponse>> GetContactsAsync(
        GetContactsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetContactsResponse>(OperationType.GetContacts, request, eventHandler, cancellationToken);

    public Task<Result<ContactDetailsDto>> GetContactDetailsAsync(
        GetContactDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<ContactDetailsDto>(OperationType.GetContactDetails, request, eventHandler, cancellationToken);

    public Task<Result> UpsertContactAsync(
        UpsertContactRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.UpsertContact, request, eventHandler, cancellationToken);

    public Task<Result> RemoveContactAsync(
        RemoveContactRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.RemoveContact, request, eventHandler, cancellationToken);

    public Task<Result<GetResourceMailboxesResponse>> GetResourceMailboxesAsync(
        GetResourceMailboxesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetResourceMailboxesResponse>(OperationType.GetResourceMailboxes, request, eventHandler, cancellationToken);

    public Task<Result<ResourceMailboxDetailsDto>> GetResourceMailboxDetailsAsync(
        GetResourceMailboxDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<ResourceMailboxDetailsDto>(OperationType.GetResourceMailboxDetails, request, eventHandler, cancellationToken);

    public Task<Result<UpsertResourceMailboxResponse>> UpsertResourceMailboxAsync(
        UpsertResourceMailboxRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<UpsertResourceMailboxResponse>(OperationType.UpsertResourceMailbox, request, eventHandler, cancellationToken);

    public Task<Result<GetPublicFoldersResponse>> GetPublicFoldersAsync(
        GetPublicFoldersRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetPublicFoldersResponse>(OperationType.GetPublicFolders, request, eventHandler, cancellationToken);

    public Task<Result<PublicFolderDetailsDto>> GetPublicFolderDetailsAsync(
        GetPublicFolderDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<PublicFolderDetailsDto>(OperationType.GetPublicFolderDetails, request, eventHandler, cancellationToken);

    public Task<Result<UpsertPublicFolderResponse>> UpsertPublicFolderAsync(
        UpsertPublicFolderRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<UpsertPublicFolderResponse>(OperationType.UpsertPublicFolder, request, eventHandler, cancellationToken);

    public Task<Result> SetPublicFolderClientPermissionAsync(
        SetPublicFolderClientPermissionRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.SetPublicFolderClientPermission, request, eventHandler, cancellationToken);

    public Task<Result> RemovePublicFolderAsync(
        RemovePublicFolderRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.RemovePublicFolder, request, eventHandler, cancellationToken);

    public Task<Result<GetMobileDevicesResponse>> GetMobileDevicesAsync(
        GetMobileDevicesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetMobileDevicesResponse>(OperationType.GetMobileDevices, request, eventHandler, cancellationToken);

    public Task<Result<GetMobileDeviceDetailsResponse>> GetMobileDeviceDetailsAsync(
        GetMobileDeviceDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetMobileDeviceDetailsResponse>(OperationType.GetMobileDeviceDetails, request, eventHandler, cancellationToken);

    public Task<Result<GetMobileDeviceMailboxPoliciesResponse>> GetMobileDeviceMailboxPoliciesAsync(
        GetMobileDeviceMailboxPoliciesRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetMobileDeviceMailboxPoliciesResponse>(
            OperationType.GetMobileDeviceMailboxPolicies,
            request ?? new GetMobileDeviceMailboxPoliciesRequest(),
            eventHandler,
            cancellationToken);

    public Task<Result> SetMobileDeviceAccessStateAsync(
        SetMobileDeviceAccessStateRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.SetMobileDeviceAccessState, request, eventHandler, cancellationToken);

    public Task<Result> ClearMobileDeviceAsync(
        ClearMobileDeviceRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.ClearMobileDevice, request, eventHandler, cancellationToken);

    public Task<Result> SetMobileDeviceMailboxPolicyAsync(
        SetMobileDeviceMailboxPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.SetMobileDeviceMailboxPolicy, request, eventHandler, cancellationToken);

    public Task<Result<GetMigrationBatchesResponse>> GetMigrationBatchesAsync(
        GetMigrationBatchesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetMigrationBatchesResponse>(OperationType.GetMigrationBatches, request, eventHandler, cancellationToken);

    public Task<Result<GetMigrationEndpointsResponse>> GetMigrationEndpointsAsync(
        GetMigrationEndpointsRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetMigrationEndpointsResponse>(
            OperationType.GetMigrationEndpoints,
            request ?? new GetMigrationEndpointsRequest(),
            eventHandler,
            cancellationToken);

    public Task<Result<MigrationBatchDetailsDto>> GetMigrationBatchDetailsAsync(
        GetMigrationBatchDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<MigrationBatchDetailsDto>(OperationType.GetMigrationBatchDetails, request, eventHandler, cancellationToken);

    public Task<Result> UpsertMigrationEndpointAsync(
        UpsertMigrationEndpointRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.UpsertMigrationEndpoint, request, eventHandler, cancellationToken);

    public Task<Result<TestMigrationEndpointResponse>> TestMigrationEndpointAsync(
        TestMigrationEndpointRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<TestMigrationEndpointResponse>(OperationType.TestMigrationEndpoint, request, eventHandler, cancellationToken);

    public Task<Result<GetMigrationBatchPreflightResponse>> GetMigrationBatchPreflightAsync(
        GetMigrationBatchPreflightRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetMigrationBatchPreflightResponse>(OperationType.GetMigrationBatchPreflight, request, eventHandler, cancellationToken);

    public Task<Result> CreateMigrationBatchAsync(
        CreateMigrationBatchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.CreateMigrationBatch, request, eventHandler, cancellationToken);

    public Task<Result> StartMigrationBatchAsync(
        StartMigrationBatchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.StartMigrationBatch, request, eventHandler, cancellationToken);

    public Task<Result> CompleteMigrationBatchAsync(
        CompleteMigrationBatchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.CompleteMigrationBatch, request, eventHandler, cancellationToken);

    public Task<Result> RemoveMigrationBatchAsync(
        RemoveMigrationBatchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.RemoveMigrationBatch, request, eventHandler, cancellationToken);

    public Task<Result<GetRoleGroupsResponse>> GetRoleGroupsAsync(
        GetRoleGroupsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<GetRoleGroupsResponse>(OperationType.GetRoleGroups, request, eventHandler, cancellationToken);

    public Task<Result<RoleGroupDetailsDto>> GetRoleGroupDetailsAsync(
        GetRoleGroupDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteOperationAsync<RoleGroupDetailsDto>(OperationType.GetRoleGroupDetails, request, eventHandler, cancellationToken);

    public Task<Result> UpsertRoleGroupAsync(
        UpsertRoleGroupRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.UpsertRoleGroup, request, eventHandler, cancellationToken);

    public Task<Result> ModifyRoleGroupMemberAsync(
        ModifyRoleGroupMemberRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _runtime.ExecuteCommandAsync(OperationType.ModifyRoleGroupMember, request, eventHandler, cancellationToken);
}
