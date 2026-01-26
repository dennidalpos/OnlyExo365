using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;
using ExchangeAdmin.Infrastructure.Ipc;

namespace ExchangeAdmin.Application.Services;

/// <summary>
/// Interface for worker service operations.
/// </summary>
public interface IWorkerService
{
    /// <summary>
    /// Current worker connection state.
    /// </summary>
    WorkerConnectionState ConnectionState { get; }

    /// <summary>
    /// Detailed worker status.
    /// </summary>
    WorkerStatus Status { get; }

    /// <summary>
    /// Cached capability map (null if not yet detected).
    /// </summary>
    CapabilityMapDto? Capabilities { get; }

    /// <summary>
    /// Connection state changed event.
    /// </summary>
    event EventHandler<WorkerConnectionState>? StateChanged;

    /// <summary>
    /// Event received from worker.
    /// </summary>
    event EventHandler<EventEnvelope>? EventReceived;

    /// <summary>
    /// Capabilities updated event.
    /// </summary>
    event EventHandler<CapabilityMapDto>? CapabilitiesUpdated;

    #region Worker Lifecycle

    /// <summary>
    /// Start the worker process.
    /// </summary>
    Task<bool> StartWorkerAsync(CancellationToken cancellationToken = default);

    /// <summary>
    /// Stop the worker process.
    /// </summary>
    Task StopWorkerAsync();

    /// <summary>
    /// Restart the worker process.
    /// </summary>
    Task<bool> RestartWorkerAsync(CancellationToken cancellationToken = default);

    /// <summary>
    /// Force kill the worker process.
    /// </summary>
    void KillWorker();

    #endregion

    #region Connection

    /// <summary>
    /// Connect to Exchange Online (interactive).
    /// </summary>
    Task<Result<ConnectionStatusDto>> ConnectExchangeAsync(
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Disconnect from Exchange Online.
    /// </summary>
    Task<Result> DisconnectExchangeAsync(CancellationToken cancellationToken = default);

    /// <summary>
    /// Get current Exchange connection status.
    /// </summary>
    Task<Result<ConnectionStatusDto>> GetConnectionStatusAsync(CancellationToken cancellationToken = default);

    /// <summary>
    /// Detect available capabilities.
    /// </summary>
    Task<Result<CapabilityMapDto>> DetectCapabilitiesAsync(
        bool forceRefresh = false,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    #endregion

    #region Dashboard

    /// <summary>
    /// Get dashboard statistics.
    /// </summary>
    Task<Result<DashboardStatsDto>> GetDashboardStatsAsync(
        GetDashboardStatsRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    #endregion

    #region Mailboxes

    /// <summary>
    /// Get mailbox list with paging.
    /// </summary>
    Task<Result<GetMailboxesResponse>> GetMailboxesAsync(
        GetMailboxesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Get mailbox details.
    /// </summary>
    Task<Result<MailboxDetailsDto>> GetMailboxDetailsAsync(
        GetMailboxDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Get retention policies.
    /// </summary>
    Task<Result<GetRetentionPoliciesResponse>> GetRetentionPoliciesAsync(
        GetRetentionPoliciesRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Set mailbox retention policy.
    /// </summary>
    Task<Result> SetRetentionPolicyAsync(
        SetRetentionPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Get mailbox permissions.
    /// </summary>
    Task<Result<MailboxPermissionsDto>> GetMailboxPermissionsAsync(
        string identity,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Set single mailbox permission.
    /// </summary>
    Task<Result> SetMailboxPermissionAsync(
        SetMailboxPermissionRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Apply permissions delta plan.
    /// </summary>
    Task<Result<ApplyPermissionsDeltaPlanResponse>> ApplyPermissionsDeltaPlanAsync(
        ApplyPermissionsDeltaPlanRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Set mailbox feature.
    /// </summary>
    Task<Result> SetMailboxFeatureAsync(
        SetMailboxFeatureRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> UpdateMailboxSettingsAsync(
        UpdateMailboxSettingsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> SetMailboxAutoReplyConfigurationAsync(
        SetMailboxAutoReplyConfigurationRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> ConvertMailboxToSharedAsync(
        ConvertMailboxToSharedRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> ConvertMailboxToRegularAsync(
        ConvertMailboxToRegularRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetMailboxSpaceReportResponse>> GetMailboxSpaceReportAsync(
        GetMailboxSpaceReportRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    #endregion

    #region Distribution Lists

    /// <summary>
    /// Get distribution lists with paging.
    /// </summary>
    Task<Result<GetDistributionListsResponse>> GetDistributionListsAsync(
        GetDistributionListsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Get distribution list details.
    /// </summary>
    Task<Result<DistributionListDetailsDto>> GetDistributionListDetailsAsync(
        GetDistributionListDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Get group members with paging.
    /// </summary>
    Task<Result<GroupMembersPageDto>> GetGroupMembersAsync(
        GetGroupMembersRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Add or remove group member.
    /// </summary>
    Task<Result> ModifyGroupMemberAsync(
        ModifyGroupMemberRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> SetDistributionListSettingsAsync(
        SetDistributionListSettingsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Preview dynamic distribution group members.
    /// </summary>
    Task<Result<PreviewDynamicGroupMembersResponse>> PreviewDynamicGroupMembersAsync(
        PreviewDynamicGroupMembersRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    #endregion

    #region Demo

    /// <summary>
    /// Run demo operation.
    /// </summary>
    Task<Result<DemoOperationResponse>> RunDemoOperationAsync(
        DemoOperationRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    #endregion

    /// <summary>
    /// Cancel an operation in progress.
    /// </summary>
    Task CancelOperationAsync(string correlationId);
}
