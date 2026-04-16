using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;

namespace ExchangeAdmin.Application.Services;

public interface IWorkerService :
    IConnectionWorkerService,
    IDashboardWorkerService,
    IResourcesWorkerService,
    IMigrationWorkerService,
    IMailboxesWorkerService,
    IDistributionListsWorkerService,
    IMailSecurityWorkerService,
    IMailFlowWorkerService,
    IComplianceWorkerService,
    ISystemWorkerService
{
    Task<Result<GetContactsResponse>> GetContactsAsync(
        GetContactsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<ContactDetailsDto>> GetContactDetailsAsync(
        GetContactDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> UpsertContactAsync(
        UpsertContactRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> RemoveContactAsync(
        RemoveContactRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetPublicFoldersResponse>> GetPublicFoldersAsync(
        GetPublicFoldersRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<PublicFolderDetailsDto>> GetPublicFolderDetailsAsync(
        GetPublicFolderDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<UpsertPublicFolderResponse>> UpsertPublicFolderAsync(
        UpsertPublicFolderRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> SetPublicFolderClientPermissionAsync(
        SetPublicFolderClientPermissionRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> RemovePublicFolderAsync(
        RemovePublicFolderRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetMobileDevicesResponse>> GetMobileDevicesAsync(
        GetMobileDevicesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetMobileDeviceDetailsResponse>> GetMobileDeviceDetailsAsync(
        GetMobileDeviceDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetMobileDeviceMailboxPoliciesResponse>> GetMobileDeviceMailboxPoliciesAsync(
        GetMobileDeviceMailboxPoliciesRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> SetMobileDeviceAccessStateAsync(
        SetMobileDeviceAccessStateRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> ClearMobileDeviceAsync(
        ClearMobileDeviceRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> SetMobileDeviceMailboxPolicyAsync(
        SetMobileDeviceMailboxPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetRoleGroupsResponse>> GetRoleGroupsAsync(
        GetRoleGroupsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<RoleGroupDetailsDto>> GetRoleGroupDetailsAsync(
        GetRoleGroupDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> UpsertRoleGroupAsync(
        UpsertRoleGroupRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> ModifyRoleGroupMemberAsync(
        ModifyRoleGroupMemberRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetMessageTraceResponse>> GetMessageTraceAsync(
        GetMessageTraceRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetMessageTraceDetailsResponse>> GetMessageTraceDetailsAsync(
        GetMessageTraceDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task CancelOperationAsync(string correlationId);
}
