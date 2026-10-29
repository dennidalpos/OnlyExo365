using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;

namespace OnlyExo365.Shell.Ipc;

public class WorkerClient : IAsyncDisposable
{
    private readonly WorkerSupervisor _supervisor;
    private readonly WorkerClientRuntime _runtime;
    private readonly WorkerConnectionClient _connectionClient;
    private readonly WorkerDashboardClient _dashboardClient;
    private readonly WorkerRecipientClient _recipientClient;
    private readonly WorkerMailboxClient _mailboxClient;
    private readonly WorkerGroupClient _groupClient;
    private readonly WorkerMailSecurityClient _mailSecurityClient;
    private readonly WorkerMailFlowClient _mailFlowClient;
    private readonly WorkerComplianceClient _complianceClient;
    private readonly WorkerSupportClient _supportClient;
    private CapabilityMapDto? _capabilities;

    public event EventHandler<WorkerConnectionState>? StateChanged;
    public event EventHandler<EventEnvelope>? EventReceived;
    public event EventHandler<CapabilityMapDto>? CapabilitiesUpdated;

    public WorkerConnectionState State => _runtime.State;
    public WorkerStatus Status => _runtime.Status;
    public CapabilityMapDto? Capabilities => _capabilities;

    public WorkerClient(WorkerSupervisorOptions? options = null)
    {
        _supervisor = new WorkerSupervisor(options);
        _runtime = new WorkerClientRuntime(_supervisor);
        _connectionClient = new WorkerConnectionClient(
            _runtime,
            () => _capabilities,
            value => _capabilities = value,
            value => CapabilitiesUpdated?.Invoke(this, value));
        _dashboardClient = new WorkerDashboardClient(_runtime);
        _recipientClient = new WorkerRecipientClient(_runtime);
        _mailboxClient = new WorkerMailboxClient(_runtime);
        _groupClient = new WorkerGroupClient(_runtime);
        _mailSecurityClient = new WorkerMailSecurityClient(_runtime);
        _mailFlowClient = new WorkerMailFlowClient(_runtime);
        _complianceClient = new WorkerComplianceClient(_runtime);
        _supportClient = new WorkerSupportClient(_runtime);

        _supervisor.StateChanged += (s, e) => StateChanged?.Invoke(this, e);
        _supervisor.EventReceived += (s, e) => EventReceived?.Invoke(this, e);
    }

    public Task<bool> StartWorkerAsync(CancellationToken cancellationToken = default)
        => _supervisor.StartAsync(cancellationToken);

    public Task StopWorkerAsync()
        => _supervisor.StopAsync();

    public Task<bool> RestartWorkerAsync(CancellationToken cancellationToken = default)
        => _supervisor.RestartAsync(cancellationToken);

    public void KillWorker()
        => _supervisor.KillWorker();

    public Task<Result<SetWorkerConsoleVisibilityResponse>> SetWorkerConsoleVisibilityAsync(
        bool isVisible,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _supportClient.SetWorkerConsoleVisibilityAsync(isVisible, eventHandler, cancellationToken);

    public Task<Result<ConnectionStatusDto>> ConnectExchangeAsync(
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _connectionClient.ConnectExchangeAsync(eventHandler, cancellationToken);

    public Task<Result> DisconnectExchangeAsync(CancellationToken cancellationToken = default)
        => _connectionClient.DisconnectExchangeAsync(cancellationToken);

    public Task<Result<ConnectionStatusDto>> GetConnectionStatusAsync(CancellationToken cancellationToken = default)
        => _connectionClient.GetConnectionStatusAsync(cancellationToken);

    public Task<Result<CapabilityMapDto>> DetectCapabilitiesAsync(
        bool forceRefresh = false,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _connectionClient.DetectCapabilitiesAsync(forceRefresh, eventHandler, cancellationToken);

    public Task<Result<DashboardStatsDto>> GetDashboardStatsAsync(
        GetDashboardStatsRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _dashboardClient.GetDashboardStatsAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetContactsResponse>> GetContactsAsync(
        GetContactsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.GetContactsAsync(request, eventHandler, cancellationToken);

    public Task<Result<ContactDetailsDto>> GetContactDetailsAsync(
        GetContactDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.GetContactDetailsAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpsertContactAsync(
        UpsertContactRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.UpsertContactAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemoveContactAsync(
        RemoveContactRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.RemoveContactAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetResourceMailboxesResponse>> GetResourceMailboxesAsync(
        GetResourceMailboxesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.GetResourceMailboxesAsync(request, eventHandler, cancellationToken);

    public Task<Result<ResourceMailboxDetailsDto>> GetResourceMailboxDetailsAsync(
        GetResourceMailboxDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.GetResourceMailboxDetailsAsync(request, eventHandler, cancellationToken);

    public Task<Result<UpsertResourceMailboxResponse>> UpsertResourceMailboxAsync(
        UpsertResourceMailboxRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.UpsertResourceMailboxAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetPublicFoldersResponse>> GetPublicFoldersAsync(
        GetPublicFoldersRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.GetPublicFoldersAsync(request, eventHandler, cancellationToken);

    public Task<Result<PublicFolderDetailsDto>> GetPublicFolderDetailsAsync(
        GetPublicFolderDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.GetPublicFolderDetailsAsync(request, eventHandler, cancellationToken);

    public Task<Result<UpsertPublicFolderResponse>> UpsertPublicFolderAsync(
        UpsertPublicFolderRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.UpsertPublicFolderAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetPublicFolderClientPermissionAsync(
        SetPublicFolderClientPermissionRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.SetPublicFolderClientPermissionAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemovePublicFolderAsync(
        RemovePublicFolderRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.RemovePublicFolderAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetMobileDevicesResponse>> GetMobileDevicesAsync(
        GetMobileDevicesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.GetMobileDevicesAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetMobileDeviceDetailsResponse>> GetMobileDeviceDetailsAsync(
        GetMobileDeviceDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.GetMobileDeviceDetailsAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetMobileDeviceMailboxPoliciesResponse>> GetMobileDeviceMailboxPoliciesAsync(
        GetMobileDeviceMailboxPoliciesRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.GetMobileDeviceMailboxPoliciesAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetMobileDeviceAccessStateAsync(
        SetMobileDeviceAccessStateRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.SetMobileDeviceAccessStateAsync(request, eventHandler, cancellationToken);

    public Task<Result> ClearMobileDeviceAsync(
        ClearMobileDeviceRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.ClearMobileDeviceAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetMobileDeviceMailboxPolicyAsync(
        SetMobileDeviceMailboxPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.SetMobileDeviceMailboxPolicyAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetMigrationBatchesResponse>> GetMigrationBatchesAsync(
        GetMigrationBatchesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.GetMigrationBatchesAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetMigrationEndpointsResponse>> GetMigrationEndpointsAsync(
        GetMigrationEndpointsRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.GetMigrationEndpointsAsync(request, eventHandler, cancellationToken);

    public Task<Result<MigrationBatchDetailsDto>> GetMigrationBatchDetailsAsync(
        GetMigrationBatchDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.GetMigrationBatchDetailsAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpsertMigrationEndpointAsync(
        UpsertMigrationEndpointRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.UpsertMigrationEndpointAsync(request, eventHandler, cancellationToken);

    public Task<Result<TestMigrationEndpointResponse>> TestMigrationEndpointAsync(
        TestMigrationEndpointRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.TestMigrationEndpointAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetMigrationBatchPreflightResponse>> GetMigrationBatchPreflightAsync(
        GetMigrationBatchPreflightRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.GetMigrationBatchPreflightAsync(request, eventHandler, cancellationToken);

    public Task<Result> CreateMigrationBatchAsync(
        CreateMigrationBatchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.CreateMigrationBatchAsync(request, eventHandler, cancellationToken);

    public Task<Result> StartMigrationBatchAsync(
        StartMigrationBatchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.StartMigrationBatchAsync(request, eventHandler, cancellationToken);

    public Task<Result> CompleteMigrationBatchAsync(
        CompleteMigrationBatchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.CompleteMigrationBatchAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemoveMigrationBatchAsync(
        RemoveMigrationBatchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.RemoveMigrationBatchAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetRoleGroupsResponse>> GetRoleGroupsAsync(
        GetRoleGroupsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.GetRoleGroupsAsync(request, eventHandler, cancellationToken);

    public Task<Result<RoleGroupDetailsDto>> GetRoleGroupDetailsAsync(
        GetRoleGroupDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.GetRoleGroupDetailsAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpsertRoleGroupAsync(
        UpsertRoleGroupRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.UpsertRoleGroupAsync(request, eventHandler, cancellationToken);

    public Task<Result> ModifyRoleGroupMemberAsync(
        ModifyRoleGroupMemberRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _recipientClient.ModifyRoleGroupMemberAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetMailboxesResponse>> GetMailboxesAsync(
        GetMailboxesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailboxClient.GetMailboxesAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetMailboxProvisioningCandidatesResponse>> GetMailboxProvisioningCandidatesAsync(
        GetMailboxProvisioningCandidatesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailboxClient.GetMailboxProvisioningCandidatesAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetDeletedMailboxesResponse>> GetDeletedMailboxesAsync(
        GetDeletedMailboxesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailboxClient.GetDeletedMailboxesAsync(request, eventHandler, cancellationToken);

    public Task<Result<MailboxDetailsDto>> GetMailboxDetailsAsync(
        GetMailboxDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailboxClient.GetMailboxDetailsAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetRetentionPoliciesResponse>> GetRetentionPoliciesAsync(
        GetRetentionPoliciesRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailboxClient.GetRetentionPoliciesAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetRetentionPolicyAsync(
        SetRetentionPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailboxClient.SetRetentionPolicyAsync(request, eventHandler, cancellationToken);

    public Task<Result<MailboxPermissionsDto>> GetMailboxPermissionsAsync(
        string identity,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailboxClient.GetMailboxPermissionsAsync(identity, eventHandler, cancellationToken);

    public Task<Result<GetMailboxFolderPermissionsResponse>> GetMailboxFolderPermissionsAsync(
        GetMailboxFolderPermissionsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailboxClient.GetMailboxFolderPermissionsAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetMailboxPermissionAsync(
        SetMailboxPermissionRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailboxClient.SetMailboxPermissionAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetMailboxFolderPermissionAsync(
        SetMailboxFolderPermissionRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailboxClient.SetMailboxFolderPermissionAsync(request, eventHandler, cancellationToken);

    public Task<Result<ApplyPermissionsDeltaPlanResponse>> ApplyPermissionsDeltaPlanAsync(
        ApplyPermissionsDeltaPlanRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailboxClient.ApplyPermissionsDeltaPlanAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpdateMailboxSettingsAsync(
        UpdateMailboxSettingsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailboxClient.UpdateMailboxSettingsAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetMailboxAutoReplyConfigurationAsync(
        SetMailboxAutoReplyConfigurationRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailboxClient.SetMailboxAutoReplyConfigurationAsync(request, eventHandler, cancellationToken);

    public Task<Result> ConvertMailboxToSharedAsync(
        ConvertMailboxToSharedRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailboxClient.ConvertMailboxToSharedAsync(request, eventHandler, cancellationToken);

    public Task<Result> ConvertMailboxToRegularAsync(
        ConvertMailboxToRegularRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailboxClient.ConvertMailboxToRegularAsync(request, eventHandler, cancellationToken);

    public Task<Result<RestoreMailboxResponse>> RestoreMailboxAsync(
        RestoreMailboxRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailboxClient.RestoreMailboxAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetMailboxSpaceReportResponse>> GetMailboxSpaceReportAsync(
        GetMailboxSpaceReportRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailboxClient.GetMailboxSpaceReportAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetMailboxAccessReportResponse>> GetMailboxAccessReportAsync(
        GetMailboxAccessReportRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailboxClient.GetMailboxAccessReportAsync(request, eventHandler, cancellationToken);

    public Task<Result> CreateMailboxAsync(
        CreateMailboxRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailboxClient.CreateMailboxAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetDistributionListsResponse>> GetDistributionListsAsync(
        GetDistributionListsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _groupClient.GetDistributionListsAsync(request, eventHandler, cancellationToken);

    public Task<Result<DistributionListDetailsDto>> GetDistributionListDetailsAsync(
        GetDistributionListDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _groupClient.GetDistributionListDetailsAsync(request, eventHandler, cancellationToken);

    public Task<Result<GroupMembersPageDto>> GetGroupMembersAsync(
        GetGroupMembersRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _groupClient.GetGroupMembersAsync(request, eventHandler, cancellationToken);

    public Task<Result> ModifyGroupMemberAsync(
        ModifyGroupMemberRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _groupClient.ModifyGroupMemberAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetDistributionListSettingsAsync(
        SetDistributionListSettingsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _groupClient.SetDistributionListSettingsAsync(request, eventHandler, cancellationToken);

    public Task<Result> CreateDistributionListAsync(
        CreateDistributionListRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _groupClient.CreateDistributionListAsync(request, eventHandler, cancellationToken);

    public Task<Result<PreviewDynamicGroupMembersResponse>> PreviewDynamicGroupMembersAsync(
        PreviewDynamicGroupMembersRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _groupClient.PreviewDynamicGroupMembersAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetMessageTraceResponse>> GetMessageTraceAsync(
        GetMessageTraceRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.GetMessageTraceAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetMessageTraceDetailsResponse>> GetMessageTraceDetailsAsync(
        GetMessageTraceDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.GetMessageTraceDetailsAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetMailSecurityBaselineResponse>> GetMailSecurityBaselineAsync(
        GetMailSecurityBaselineRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailSecurityClient.GetMailSecurityBaselineAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpdateDkimSigningConfigAsync(
        UpdateDkimSigningConfigRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailSecurityClient.UpdateDkimSigningConfigAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpdateHostedContentFilterPolicyAsync(
        UpdateHostedContentFilterPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailSecurityClient.UpdateHostedContentFilterPolicyAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpdateAntiPhishPolicyAsync(
        UpdateAntiPhishPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailSecurityClient.UpdateAntiPhishPolicyAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpdateMalwareFilterPolicyAsync(
        UpdateMalwareFilterPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailSecurityClient.UpdateMalwareFilterPolicyAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpdateQuarantinePolicyAsync(
        UpdateQuarantinePolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailSecurityClient.UpdateQuarantinePolicyAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpdateHostedOutboundSpamFilterPolicyAsync(
        UpdateHostedOutboundSpamFilterPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailSecurityClient.UpdateHostedOutboundSpamFilterPolicyAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetComplianceWorkspaceResponse>> GetComplianceWorkspaceAsync(
        GetComplianceWorkspaceRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _complianceClient.GetComplianceWorkspaceAsync(request, eventHandler, cancellationToken);

    public Task<Result<SearchUnifiedAuditLogResponse>> SearchUnifiedAuditLogAsync(
        SearchUnifiedAuditLogRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _complianceClient.SearchUnifiedAuditLogAsync(request, eventHandler, cancellationToken);

    public Task<Result> CreateComplianceSearchAsync(
        CreateComplianceSearchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _complianceClient.CreateComplianceSearchAsync(request, eventHandler, cancellationToken);

    public Task<Result> StartComplianceSearchAsync(
        StartComplianceSearchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _complianceClient.StartComplianceSearchAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemoveComplianceSearchAsync(
        RemoveComplianceSearchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _complianceClient.RemoveComplianceSearchAsync(request, eventHandler, cancellationToken);

    public Task<Result<InvokeComplianceActionResponse>> InvokeComplianceActionAsync(
        InvokeComplianceActionRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _complianceClient.InvokeComplianceActionAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetTransportRulesResponse>> GetTransportRulesAsync(
        GetTransportRulesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.GetTransportRulesAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetTransportRuleStateAsync(
        SetTransportRuleStateRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.SetTransportRuleStateAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpsertTransportRuleAsync(
        UpsertTransportRuleRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.UpsertTransportRuleAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemoveTransportRuleAsync(
        RemoveTransportRuleRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.RemoveTransportRuleAsync(request, eventHandler, cancellationToken);

    public Task<Result<TestTransportRuleResponse>> TestTransportRuleAsync(
        TestTransportRuleRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.TestTransportRuleAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetConnectorsResponse>> GetConnectorsAsync(
        GetConnectorsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.GetConnectorsAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetAcceptedDomainsResponse>> GetAcceptedDomainsAsync(
        GetAcceptedDomainsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.GetAcceptedDomainsAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpsertConnectorAsync(
        UpsertConnectorRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.UpsertConnectorAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemoveConnectorAsync(
        RemoveConnectorRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.RemoveConnectorAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpsertAcceptedDomainAsync(
        UpsertAcceptedDomainRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.UpsertAcceptedDomainAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemoveAcceptedDomainAsync(
        RemoveAcceptedDomainRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.RemoveAcceptedDomainAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetRemoteDomainsResponse>> GetRemoteDomainsAsync(
        GetRemoteDomainsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.GetRemoteDomainsAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpsertRemoteDomainAsync(
        UpsertRemoteDomainRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.UpsertRemoteDomainAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemoveRemoteDomainAsync(
        RemoveRemoteDomainRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.RemoveRemoteDomainAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetOrganizationRelationshipsResponse>> GetOrganizationRelationshipsAsync(
        GetOrganizationRelationshipsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.GetOrganizationRelationshipsAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetAddressListsResponse>> GetAddressListsAsync(
        GetAddressListsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.GetAddressListsAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpsertAddressListAsync(
        UpsertAddressListRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.UpsertAddressListAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemoveAddressListAsync(
        RemoveAddressListRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.RemoveAddressListAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetAddressBookPoliciesResponse>> GetAddressBookPoliciesAsync(
        GetAddressBookPoliciesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.GetAddressBookPoliciesAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpsertAddressBookPolicyAsync(
        UpsertAddressBookPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.UpsertAddressBookPolicyAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemoveAddressBookPolicyAsync(
        RemoveAddressBookPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.RemoveAddressBookPolicyAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetOfflineAddressBooksResponse>> GetOfflineAddressBooksAsync(
        GetOfflineAddressBooksRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.GetOfflineAddressBooksAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpsertOfflineAddressBookAsync(
        UpsertOfflineAddressBookRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.UpsertOfflineAddressBookAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemoveOfflineAddressBookAsync(
        RemoveOfflineAddressBookRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.RemoveOfflineAddressBookAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetSharingPoliciesResponse>> GetSharingPoliciesAsync(
        GetSharingPoliciesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.GetSharingPoliciesAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpsertSharingPolicyAsync(
        UpsertSharingPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.UpsertSharingPolicyAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemoveSharingPolicyAsync(
        RemoveSharingPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.RemoveSharingPolicyAsync(request, eventHandler, cancellationToken);

    public Task<Result> UpsertOrganizationRelationshipAsync(
        UpsertOrganizationRelationshipRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.UpsertOrganizationRelationshipAsync(request, eventHandler, cancellationToken);

    public Task<Result> RemoveOrganizationRelationshipAsync(
        RemoveOrganizationRelationshipRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _mailFlowClient.RemoveOrganizationRelationshipAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetUserLicensesResponse>> GetUserLicensesAsync(
        GetUserLicensesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _supportClient.GetUserLicensesAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetUserLicenseAsync(
        SetUserLicenseRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _supportClient.SetUserLicenseAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetUsageLocationSuggestionResponse>> GetUsageLocationSuggestionAsync(
        GetUsageLocationSuggestionRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _supportClient.GetUsageLocationSuggestionAsync(request, eventHandler, cancellationToken);

    public Task<Result> SetUserUsageLocationAsync(
        SetUserUsageLocationRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _supportClient.SetUserUsageLocationAsync(request, eventHandler, cancellationToken);

    public Task<Result<GetAvailableLicensesResponse>> GetAvailableLicensesAsync(
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _supportClient.GetAvailableLicensesAsync(eventHandler, cancellationToken);

    public Task<Result<PrerequisiteStatusDto>> CheckPrerequisitesAsync(
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _supportClient.CheckPrerequisitesAsync(eventHandler, cancellationToken);

    public Task<Result<InstallModuleResponse>> InstallModuleAsync(
        InstallModuleRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default)
        => _supportClient.InstallModuleAsync(request, eventHandler, cancellationToken);

    public Task CancelOperationAsync(string correlationId)
        => _runtime.CancelOperationAsync(correlationId);

    public async ValueTask DisposeAsync()
    {
        await _supervisor.DisposeAsync();
    }
}

