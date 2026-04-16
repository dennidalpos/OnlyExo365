using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;
using ExchangeAdmin.Infrastructure.Ipc;

namespace ExchangeAdmin.Application.Services;

             
                                  
              
public class WorkerService : IWorkerService, IAsyncDisposable
{
    private readonly WorkerClient _client;

    public WorkerConnectionState ConnectionState => _client.State;
    public WorkerStatus Status => _client.Status;
    public CapabilityMapDto? Capabilities => _client.Capabilities;

    public event EventHandler<WorkerConnectionState>? StateChanged;
    public event EventHandler<EventEnvelope>? EventReceived;
    public event EventHandler<CapabilityMapDto>? CapabilitiesUpdated;

    public WorkerService(WorkerSupervisorOptions? options = null)
    {
        _client = new WorkerClient(options);
        _client.StateChanged += (s, e) => StateChanged?.Invoke(this, e);
        _client.EventReceived += (s, e) => EventReceived?.Invoke(this, e);
        _client.CapabilitiesUpdated += (s, e) => CapabilitiesUpdated?.Invoke(this, e);
    }

    #region Worker Lifecycle

    public Task<bool> StartWorkerAsync(CancellationToken cancellationToken = default)
        => _client.StartWorkerAsync(cancellationToken);

    public Task StopWorkerAsync()
        => _client.StopWorkerAsync();

    public Task<bool> RestartWorkerAsync(CancellationToken cancellationToken = default)
        => _client.RestartWorkerAsync(cancellationToken);

    public void KillWorker()
        => _client.KillWorker();

    public Task<Result<SetWorkerConsoleVisibilityResponse>> SetWorkerConsoleVisibilityAsync(
        bool isVisible,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.SetWorkerConsoleVisibilityAsync(isVisible, eventHandler, cancellationToken);

    #endregion

    #region Connection

    public Task<Result<ConnectionStatusDto>> ConnectExchangeAsync(
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.ConnectExchangeAsync(eventHandler, cancellationToken);

    public Task<Result> DisconnectExchangeAsync(CancellationToken cancellationToken = default)
        => _client.DisconnectExchangeAsync(cancellationToken);

    public Task<Result<ConnectionStatusDto>> GetConnectionStatusAsync(CancellationToken cancellationToken = default)
        => _client.GetConnectionStatusAsync(cancellationToken);

    public Task<Result<CapabilityMapDto>> DetectCapabilitiesAsync(
        bool forceRefresh = false,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.DetectCapabilitiesAsync(forceRefresh, eventHandler, cancellationToken);

    #endregion

    #region Dashboard

    public Task<Result<DashboardStatsDto>> GetDashboardStatsAsync(
        GetDashboardStatsRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetDashboardStatsAsync(request, eventHandler, cancellationToken);

    #endregion

    #region Contacts

    public Task<Result<GetContactsResponse>> GetContactsAsync(
        GetContactsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetContactsAsync(request, eventHandler, cancellationToken);

    public Task<Result<ContactDetailsDto>> GetContactDetailsAsync(
        GetContactDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetContactDetailsAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpsertContactAsync(
        UpsertContactRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.UpsertContactAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemoveContactAsync(
        RemoveContactRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.RemoveContactAsync(request, eventHandler, cancellationToken);

    #endregion

    #region Resources

    public Task<Result<GetResourceMailboxesResponse>> GetResourceMailboxesAsync(
        GetResourceMailboxesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetResourceMailboxesAsync(request, eventHandler, cancellationToken);

    public Task<Result<ResourceMailboxDetailsDto>> GetResourceMailboxDetailsAsync(
        GetResourceMailboxDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetResourceMailboxDetailsAsync(request, eventHandler, cancellationToken);

    public Task<Result<UpsertResourceMailboxResponse>> UpsertResourceMailboxAsync(
        UpsertResourceMailboxRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.UpsertResourceMailboxAsync(request, eventHandler, cancellationToken);

    #endregion

    #region Public Folders

    public Task<Result<GetPublicFoldersResponse>> GetPublicFoldersAsync(
        GetPublicFoldersRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetPublicFoldersAsync(request, eventHandler, cancellationToken);

    public Task<Result<PublicFolderDetailsDto>> GetPublicFolderDetailsAsync(
        GetPublicFolderDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetPublicFolderDetailsAsync(request, eventHandler, cancellationToken);

    public Task<Result<UpsertPublicFolderResponse>> UpsertPublicFolderAsync(
        UpsertPublicFolderRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.UpsertPublicFolderAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetPublicFolderClientPermissionAsync(
        SetPublicFolderClientPermissionRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.SetPublicFolderClientPermissionAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemovePublicFolderAsync(
        RemovePublicFolderRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.RemovePublicFolderAsync(request, eventHandler, cancellationToken);

    #endregion

    #region Mobile Devices

    public Task<Result<GetMobileDevicesResponse>> GetMobileDevicesAsync(
        GetMobileDevicesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetMobileDevicesAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetMobileDeviceDetailsResponse>> GetMobileDeviceDetailsAsync(
        GetMobileDeviceDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetMobileDeviceDetailsAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetMobileDeviceMailboxPoliciesResponse>> GetMobileDeviceMailboxPoliciesAsync(
        GetMobileDeviceMailboxPoliciesRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetMobileDeviceMailboxPoliciesAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetMobileDeviceAccessStateAsync(
        SetMobileDeviceAccessStateRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.SetMobileDeviceAccessStateAsync(request, eventHandler, cancellationToken);

    public Task<Result> ClearMobileDeviceAsync(
        ClearMobileDeviceRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.ClearMobileDeviceAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetMobileDeviceMailboxPolicyAsync(
        SetMobileDeviceMailboxPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.SetMobileDeviceMailboxPolicyAsync(request, eventHandler, cancellationToken);

    #endregion

    #region Migration

    public Task<Result<GetMigrationBatchesResponse>> GetMigrationBatchesAsync(
        GetMigrationBatchesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetMigrationBatchesAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetMigrationEndpointsResponse>> GetMigrationEndpointsAsync(
        GetMigrationEndpointsRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetMigrationEndpointsAsync(request, eventHandler, cancellationToken);

    public Task<Result<MigrationBatchDetailsDto>> GetMigrationBatchDetailsAsync(
        GetMigrationBatchDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetMigrationBatchDetailsAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpsertMigrationEndpointAsync(
        UpsertMigrationEndpointRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.UpsertMigrationEndpointAsync(request, eventHandler, cancellationToken);

    public Task<Result<TestMigrationEndpointResponse>> TestMigrationEndpointAsync(
        TestMigrationEndpointRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.TestMigrationEndpointAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetMigrationBatchPreflightResponse>> GetMigrationBatchPreflightAsync(
        GetMigrationBatchPreflightRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetMigrationBatchPreflightAsync(request, eventHandler, cancellationToken);

    public Task<Result> CreateMigrationBatchAsync(
        CreateMigrationBatchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.CreateMigrationBatchAsync(request, eventHandler, cancellationToken);

    public Task<Result> StartMigrationBatchAsync(
        StartMigrationBatchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.StartMigrationBatchAsync(request, eventHandler, cancellationToken);

    public Task<Result> CompleteMigrationBatchAsync(
        CompleteMigrationBatchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.CompleteMigrationBatchAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemoveMigrationBatchAsync(
        RemoveMigrationBatchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.RemoveMigrationBatchAsync(request, eventHandler, cancellationToken);

    #endregion

    #region Permissions

    public Task<Result<GetRoleGroupsResponse>> GetRoleGroupsAsync(
        GetRoleGroupsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetRoleGroupsAsync(request, eventHandler, cancellationToken);

    public Task<Result<RoleGroupDetailsDto>> GetRoleGroupDetailsAsync(
        GetRoleGroupDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetRoleGroupDetailsAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpsertRoleGroupAsync(
        UpsertRoleGroupRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.UpsertRoleGroupAsync(request, eventHandler, cancellationToken);

    public Task<Result> ModifyRoleGroupMemberAsync(
        ModifyRoleGroupMemberRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.ModifyRoleGroupMemberAsync(request, eventHandler, cancellationToken);

    #endregion

    #region Mailboxes

    public Task<Result<GetMailboxProvisioningCandidatesResponse>> GetMailboxProvisioningCandidatesAsync(
        GetMailboxProvisioningCandidatesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetMailboxProvisioningCandidatesAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetMailboxesResponse>> GetMailboxesAsync(
        GetMailboxesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetMailboxesAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetDeletedMailboxesResponse>> GetDeletedMailboxesAsync(
        GetDeletedMailboxesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetDeletedMailboxesAsync(request, eventHandler, cancellationToken);

    public Task<Result<MailboxDetailsDto>> GetMailboxDetailsAsync(
        GetMailboxDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetMailboxDetailsAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetRetentionPoliciesResponse>> GetRetentionPoliciesAsync(
        GetRetentionPoliciesRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetRetentionPoliciesAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetRetentionPolicyAsync(
        SetRetentionPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.SetRetentionPolicyAsync(request, eventHandler, cancellationToken);

    public Task<Result<MailboxPermissionsDto>> GetMailboxPermissionsAsync(
        string identity,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetMailboxPermissionsAsync(identity, eventHandler, cancellationToken);

    public Task<Result<GetMailboxFolderPermissionsResponse>> GetMailboxFolderPermissionsAsync(
        GetMailboxFolderPermissionsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetMailboxFolderPermissionsAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetMailboxPermissionAsync(
        SetMailboxPermissionRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.SetMailboxPermissionAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetMailboxFolderPermissionAsync(
        SetMailboxFolderPermissionRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.SetMailboxFolderPermissionAsync(request, eventHandler, cancellationToken);

    public Task<Result<ApplyPermissionsDeltaPlanResponse>> ApplyPermissionsDeltaPlanAsync(
        ApplyPermissionsDeltaPlanRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.ApplyPermissionsDeltaPlanAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpdateMailboxSettingsAsync(
        UpdateMailboxSettingsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.UpdateMailboxSettingsAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetMailboxAutoReplyConfigurationAsync(
        SetMailboxAutoReplyConfigurationRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.SetMailboxAutoReplyConfigurationAsync(request, eventHandler, cancellationToken);

    public Task<Result> ConvertMailboxToSharedAsync(
        ConvertMailboxToSharedRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.ConvertMailboxToSharedAsync(request, eventHandler, cancellationToken);

    public Task<Result> ConvertMailboxToRegularAsync(
        ConvertMailboxToRegularRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.ConvertMailboxToRegularAsync(request, eventHandler, cancellationToken);

    public Task<Result<RestoreMailboxResponse>> RestoreMailboxAsync(
        RestoreMailboxRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.RestoreMailboxAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetMailboxSpaceReportResponse>> GetMailboxSpaceReportAsync(
        GetMailboxSpaceReportRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetMailboxSpaceReportAsync(request, eventHandler, cancellationToken);


    public Task<Result<GetMailboxAccessReportResponse>> GetMailboxAccessReportAsync(
        GetMailboxAccessReportRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetMailboxAccessReportAsync(request, eventHandler, cancellationToken);

    public Task<Result> CreateMailboxAsync(
        CreateMailboxRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.CreateMailboxAsync(request, eventHandler, cancellationToken);

    #endregion

    #region Distribution Lists

    public Task<Result<GetDistributionListsResponse>> GetDistributionListsAsync(
        GetDistributionListsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetDistributionListsAsync(request, eventHandler, cancellationToken);

    public Task<Result<DistributionListDetailsDto>> GetDistributionListDetailsAsync(
        GetDistributionListDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetDistributionListDetailsAsync(request, eventHandler, cancellationToken);

    public Task<Result<GroupMembersPageDto>> GetGroupMembersAsync(
        GetGroupMembersRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetGroupMembersAsync(request, eventHandler, cancellationToken);

    public Task<Result> ModifyGroupMemberAsync(
        ModifyGroupMemberRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.ModifyGroupMemberAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetDistributionListSettingsAsync(
        SetDistributionListSettingsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.SetDistributionListSettingsAsync(request, eventHandler, cancellationToken);

    public Task<Result> CreateDistributionListAsync(
        CreateDistributionListRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.CreateDistributionListAsync(request, eventHandler, cancellationToken);

    public Task<Result<PreviewDynamicGroupMembersResponse>> PreviewDynamicGroupMembersAsync(
        PreviewDynamicGroupMembersRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.PreviewDynamicGroupMembersAsync(request, eventHandler, cancellationToken);

    #endregion

    #region Message Trace

    public Task<Result<GetMessageTraceResponse>> GetMessageTraceAsync(
        GetMessageTraceRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetMessageTraceAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetMessageTraceDetailsResponse>> GetMessageTraceDetailsAsync(
        GetMessageTraceDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetMessageTraceDetailsAsync(request, eventHandler, cancellationToken);

    #endregion

    #region Mail Security

    public Task<Result<GetMailSecurityBaselineResponse>> GetMailSecurityBaselineAsync(
        GetMailSecurityBaselineRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetMailSecurityBaselineAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpdateDkimSigningConfigAsync(
        UpdateDkimSigningConfigRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.UpdateDkimSigningConfigAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpdateHostedContentFilterPolicyAsync(
        UpdateHostedContentFilterPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.UpdateHostedContentFilterPolicyAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpdateAntiPhishPolicyAsync(
        UpdateAntiPhishPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.UpdateAntiPhishPolicyAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpdateMalwareFilterPolicyAsync(
        UpdateMalwareFilterPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.UpdateMalwareFilterPolicyAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpdateQuarantinePolicyAsync(
        UpdateQuarantinePolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.UpdateQuarantinePolicyAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpdateHostedOutboundSpamFilterPolicyAsync(
        UpdateHostedOutboundSpamFilterPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.UpdateHostedOutboundSpamFilterPolicyAsync(request, eventHandler, cancellationToken);

    #endregion

    #region Compliance

    public Task<Result<GetComplianceWorkspaceResponse>> GetComplianceWorkspaceAsync(
        GetComplianceWorkspaceRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetComplianceWorkspaceAsync(request, eventHandler, cancellationToken);

    public Task<Result<SearchUnifiedAuditLogResponse>> SearchUnifiedAuditLogAsync(
        SearchUnifiedAuditLogRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.SearchUnifiedAuditLogAsync(request, eventHandler, cancellationToken);

    public Task<Result> CreateComplianceSearchAsync(
        CreateComplianceSearchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.CreateComplianceSearchAsync(request, eventHandler, cancellationToken);

    public Task<Result> StartComplianceSearchAsync(
        StartComplianceSearchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.StartComplianceSearchAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemoveComplianceSearchAsync(
        RemoveComplianceSearchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.RemoveComplianceSearchAsync(request, eventHandler, cancellationToken);

    public Task<Result<InvokeComplianceActionResponse>> InvokeComplianceActionAsync(
        InvokeComplianceActionRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.InvokeComplianceActionAsync(request, eventHandler, cancellationToken);

    #endregion

    #region Mail Flow

    public Task<Result<GetTransportRulesResponse>> GetTransportRulesAsync(
        GetTransportRulesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetTransportRulesAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetTransportRuleStateAsync(
        SetTransportRuleStateRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.SetTransportRuleStateAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpsertTransportRuleAsync(
        UpsertTransportRuleRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.UpsertTransportRuleAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemoveTransportRuleAsync(
        RemoveTransportRuleRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.RemoveTransportRuleAsync(request, eventHandler, cancellationToken);

    public Task<Result<TestTransportRuleResponse>> TestTransportRuleAsync(
        TestTransportRuleRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.TestTransportRuleAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetConnectorsResponse>> GetConnectorsAsync(
        GetConnectorsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetConnectorsAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetAcceptedDomainsResponse>> GetAcceptedDomainsAsync(
        GetAcceptedDomainsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetAcceptedDomainsAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpsertConnectorAsync(
        UpsertConnectorRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.UpsertConnectorAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemoveConnectorAsync(
        RemoveConnectorRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.RemoveConnectorAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpsertAcceptedDomainAsync(
        UpsertAcceptedDomainRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.UpsertAcceptedDomainAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemoveAcceptedDomainAsync(
        RemoveAcceptedDomainRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.RemoveAcceptedDomainAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetRemoteDomainsResponse>> GetRemoteDomainsAsync(
        GetRemoteDomainsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetRemoteDomainsAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpsertRemoteDomainAsync(
        UpsertRemoteDomainRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.UpsertRemoteDomainAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemoveRemoteDomainAsync(
        RemoveRemoteDomainRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.RemoveRemoteDomainAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetOrganizationRelationshipsResponse>> GetOrganizationRelationshipsAsync(
        GetOrganizationRelationshipsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetOrganizationRelationshipsAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetAddressListsResponse>> GetAddressListsAsync(
        GetAddressListsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetAddressListsAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpsertAddressListAsync(
        UpsertAddressListRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.UpsertAddressListAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemoveAddressListAsync(
        RemoveAddressListRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.RemoveAddressListAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetAddressBookPoliciesResponse>> GetAddressBookPoliciesAsync(
        GetAddressBookPoliciesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetAddressBookPoliciesAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpsertAddressBookPolicyAsync(
        UpsertAddressBookPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.UpsertAddressBookPolicyAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemoveAddressBookPolicyAsync(
        RemoveAddressBookPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.RemoveAddressBookPolicyAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetOfflineAddressBooksResponse>> GetOfflineAddressBooksAsync(
        GetOfflineAddressBooksRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetOfflineAddressBooksAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpsertOfflineAddressBookAsync(
        UpsertOfflineAddressBookRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.UpsertOfflineAddressBookAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemoveOfflineAddressBookAsync(
        RemoveOfflineAddressBookRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.RemoveOfflineAddressBookAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetSharingPoliciesResponse>> GetSharingPoliciesAsync(
        GetSharingPoliciesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetSharingPoliciesAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpsertSharingPolicyAsync(
        UpsertSharingPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.UpsertSharingPolicyAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemoveSharingPolicyAsync(
        RemoveSharingPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.RemoveSharingPolicyAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpsertOrganizationRelationshipAsync(
        UpsertOrganizationRelationshipRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.UpsertOrganizationRelationshipAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemoveOrganizationRelationshipAsync(
        RemoveOrganizationRelationshipRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.RemoveOrganizationRelationshipAsync(request, eventHandler, cancellationToken);

    #endregion

    #region Licenses

    public Task<Result<GetUserLicensesResponse>> GetUserLicensesAsync(
        GetUserLicensesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetUserLicensesAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetUserLicenseAsync(
        SetUserLicenseRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.SetUserLicenseAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetUsageLocationSuggestionResponse>> GetUsageLocationSuggestionAsync(
        GetUsageLocationSuggestionRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetUsageLocationSuggestionAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetUserUsageLocationAsync(
        SetUserUsageLocationRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.SetUserUsageLocationAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetAvailableLicensesResponse>> GetAvailableLicensesAsync(
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.GetAvailableLicensesAsync(eventHandler, cancellationToken);

    #endregion

    #region System

    public Task<Result<PrerequisiteStatusDto>> CheckPrerequisitesAsync(
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.CheckPrerequisitesAsync(eventHandler, cancellationToken);

    public Task<Result<InstallModuleResponse>> InstallModuleAsync(
        InstallModuleRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _client.InstallModuleAsync(request, eventHandler, cancellationToken);

    #endregion

    public Task CancelOperationAsync(string correlationId)
        => _client.CancelOperationAsync(correlationId);

    public async ValueTask DisposeAsync()
    {
        await _client.DisposeAsync();
    }
}
