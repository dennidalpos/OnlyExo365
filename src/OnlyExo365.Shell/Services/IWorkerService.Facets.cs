using OnlyExo365.Contracts.Dtos;
using OnlyExo365.Contracts.Messages;
using OnlyExo365.Contracts.Results;
using OnlyExo365.Shell.Ipc;
using OnlyExo365.Contracts.Diagnostics;

namespace OnlyExo365.Shell.Services;

public interface IInteractiveExchangeBootstrapService
{
    Task<Result> EnsureReadyAsync(
        Action<LogLevel, string>? onLog = null,
        CancellationToken cancellationToken = default);
}

public interface IConnectionWorkerService
{
    WorkerConnectionState ConnectionState { get; }
    WorkerStatus Status { get; }
    CapabilityMapDto? Capabilities { get; }

    event EventHandler<WorkerConnectionState>? StateChanged;
    event EventHandler<EventEnvelope>? EventReceived;
    event EventHandler<CapabilityMapDto>? CapabilitiesUpdated;

    Task<bool> StartWorkerAsync(CancellationToken cancellationToken = default);
    Task StopWorkerAsync();
    Task<bool> RestartWorkerAsync(CancellationToken cancellationToken = default);
    void KillWorker();
    Task<Result<SetWorkerConsoleVisibilityResponse>> SetWorkerConsoleVisibilityAsync(
        bool isVisible,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<ConnectionStatusDto>> ConnectExchangeAsync(
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> DisconnectExchangeAsync(CancellationToken cancellationToken = default);

    Task<Result<ConnectionStatusDto>> GetConnectionStatusAsync(CancellationToken cancellationToken = default);

    Task<Result<CapabilityMapDto>> DetectCapabilitiesAsync(
        bool forceRefresh = false,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);
}

public interface IDashboardWorkerService
{
    Task<Result<DashboardStatsDto>> GetDashboardStatsAsync(
        GetDashboardStatsRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);
}

public interface IResourcesWorkerService
{
    Task<Result<GetResourceMailboxesResponse>> GetResourceMailboxesAsync(
        GetResourceMailboxesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<ResourceMailboxDetailsDto>> GetResourceMailboxDetailsAsync(
        GetResourceMailboxDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<UpsertResourceMailboxResponse>> UpsertResourceMailboxAsync(
        UpsertResourceMailboxRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<ApplyPermissionsDeltaPlanResponse>> ApplyPermissionsDeltaPlanAsync(
        ApplyPermissionsDeltaPlanRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);
}

public interface IMigrationWorkerService
{
    Task<Result<GetMigrationBatchesResponse>> GetMigrationBatchesAsync(
        GetMigrationBatchesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetMigrationEndpointsResponse>> GetMigrationEndpointsAsync(
        GetMigrationEndpointsRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<MigrationBatchDetailsDto>> GetMigrationBatchDetailsAsync(
        GetMigrationBatchDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> UpsertMigrationEndpointAsync(
        UpsertMigrationEndpointRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<TestMigrationEndpointResponse>> TestMigrationEndpointAsync(
        TestMigrationEndpointRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetMigrationBatchPreflightResponse>> GetMigrationBatchPreflightAsync(
        GetMigrationBatchPreflightRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> CreateMigrationBatchAsync(
        CreateMigrationBatchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> StartMigrationBatchAsync(
        StartMigrationBatchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> CompleteMigrationBatchAsync(
        CompleteMigrationBatchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> RemoveMigrationBatchAsync(
        RemoveMigrationBatchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);
}

public interface IMailboxesWorkerService
{
    Task<Result<GetMailboxProvisioningCandidatesResponse>> GetMailboxProvisioningCandidatesAsync(
        GetMailboxProvisioningCandidatesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetMailboxesResponse>> GetMailboxesAsync(
        GetMailboxesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetDeletedMailboxesResponse>> GetDeletedMailboxesAsync(
        GetDeletedMailboxesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<MailboxDetailsDto>> GetMailboxDetailsAsync(
        GetMailboxDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetRetentionPoliciesResponse>> GetRetentionPoliciesAsync(
        GetRetentionPoliciesRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> SetRetentionPolicyAsync(
        SetRetentionPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<ApplyPermissionsDeltaPlanResponse>> ApplyPermissionsDeltaPlanAsync(
        ApplyPermissionsDeltaPlanRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetMailboxFolderPermissionsResponse>> GetMailboxFolderPermissionsAsync(
        GetMailboxFolderPermissionsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> SetMailboxFolderPermissionAsync(
        SetMailboxFolderPermissionRequest request,
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

    Task<Result<RestoreMailboxResponse>> RestoreMailboxAsync(
        RestoreMailboxRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetMailboxSpaceReportResponse>> GetMailboxSpaceReportAsync(
        GetMailboxSpaceReportRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetMailboxAccessReportResponse>> GetMailboxAccessReportAsync(
        GetMailboxAccessReportRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> CreateMailboxAsync(
        CreateMailboxRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetAcceptedDomainsResponse>> GetAcceptedDomainsAsync(
        GetAcceptedDomainsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetUserLicensesResponse>> GetUserLicensesAsync(
        GetUserLicensesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> SetUserLicenseAsync(
        SetUserLicenseRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetUsageLocationSuggestionResponse>> GetUsageLocationSuggestionAsync(
        GetUsageLocationSuggestionRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> SetUserUsageLocationAsync(
        SetUserUsageLocationRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetAvailableLicensesResponse>> GetAvailableLicensesAsync(
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);
}

public interface IDistributionListsWorkerService
{
    Task<Result<GetDistributionListsResponse>> GetDistributionListsAsync(
        GetDistributionListsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<DistributionListDetailsDto>> GetDistributionListDetailsAsync(
        GetDistributionListDetailsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GroupMembersPageDto>> GetGroupMembersAsync(
        GetGroupMembersRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> ModifyGroupMemberAsync(
        ModifyGroupMemberRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> SetDistributionListSettingsAsync(
        SetDistributionListSettingsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> CreateDistributionListAsync(
        CreateDistributionListRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<PreviewDynamicGroupMembersResponse>> PreviewDynamicGroupMembersAsync(
        PreviewDynamicGroupMembersRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetAcceptedDomainsResponse>> GetAcceptedDomainsAsync(
        GetAcceptedDomainsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);
}

public interface IMailFlowWorkerService
{
    Task<Result<GetTransportRulesResponse>> GetTransportRulesAsync(
        GetTransportRulesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> SetTransportRuleStateAsync(
        SetTransportRuleStateRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> UpsertTransportRuleAsync(
        UpsertTransportRuleRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> RemoveTransportRuleAsync(
        RemoveTransportRuleRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<TestTransportRuleResponse>> TestTransportRuleAsync(
        TestTransportRuleRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetConnectorsResponse>> GetConnectorsAsync(
        GetConnectorsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetAcceptedDomainsResponse>> GetAcceptedDomainsAsync(
        GetAcceptedDomainsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> UpsertConnectorAsync(
        UpsertConnectorRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> RemoveConnectorAsync(
        RemoveConnectorRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> UpsertAcceptedDomainAsync(
        UpsertAcceptedDomainRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> RemoveAcceptedDomainAsync(
        RemoveAcceptedDomainRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetRemoteDomainsResponse>> GetRemoteDomainsAsync(
        GetRemoteDomainsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> UpsertRemoteDomainAsync(
        UpsertRemoteDomainRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> RemoveRemoteDomainAsync(
        RemoveRemoteDomainRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetOrganizationRelationshipsResponse>> GetOrganizationRelationshipsAsync(
        GetOrganizationRelationshipsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetAddressListsResponse>> GetAddressListsAsync(
        GetAddressListsRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> UpsertAddressListAsync(
        UpsertAddressListRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> RemoveAddressListAsync(
        RemoveAddressListRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetAddressBookPoliciesResponse>> GetAddressBookPoliciesAsync(
        GetAddressBookPoliciesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> UpsertAddressBookPolicyAsync(
        UpsertAddressBookPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> RemoveAddressBookPolicyAsync(
        RemoveAddressBookPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetOfflineAddressBooksResponse>> GetOfflineAddressBooksAsync(
        GetOfflineAddressBooksRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> UpsertOfflineAddressBookAsync(
        UpsertOfflineAddressBookRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> RemoveOfflineAddressBookAsync(
        RemoveOfflineAddressBookRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<GetSharingPoliciesResponse>> GetSharingPoliciesAsync(
        GetSharingPoliciesRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> UpsertSharingPolicyAsync(
        UpsertSharingPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> RemoveSharingPolicyAsync(
        RemoveSharingPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> UpsertOrganizationRelationshipAsync(
        UpsertOrganizationRelationshipRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> RemoveOrganizationRelationshipAsync(
        RemoveOrganizationRelationshipRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);
}

public interface IMailSecurityWorkerService
{
    Task<Result<GetMailSecurityBaselineResponse>> GetMailSecurityBaselineAsync(
        GetMailSecurityBaselineRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> UpdateDkimSigningConfigAsync(
        UpdateDkimSigningConfigRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> UpdateHostedContentFilterPolicyAsync(
        UpdateHostedContentFilterPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> UpdateAntiPhishPolicyAsync(
        UpdateAntiPhishPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> UpdateMalwareFilterPolicyAsync(
        UpdateMalwareFilterPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> UpdateQuarantinePolicyAsync(
        UpdateQuarantinePolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> UpdateHostedOutboundSpamFilterPolicyAsync(
        UpdateHostedOutboundSpamFilterPolicyRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);
}

public interface IComplianceWorkerService
{
    Task<Result<GetComplianceWorkspaceResponse>> GetComplianceWorkspaceAsync(
        GetComplianceWorkspaceRequest? request = null,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<SearchUnifiedAuditLogResponse>> SearchUnifiedAuditLogAsync(
        SearchUnifiedAuditLogRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> CreateComplianceSearchAsync(
        CreateComplianceSearchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> StartComplianceSearchAsync(
        StartComplianceSearchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result> RemoveComplianceSearchAsync(
        RemoveComplianceSearchRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<InvokeComplianceActionResponse>> InvokeComplianceActionAsync(
        InvokeComplianceActionRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);
}

public interface ISystemWorkerService
{
    Task<Result<PrerequisiteStatusDto>> CheckPrerequisitesAsync(
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);

    Task<Result<InstallModuleResponse>> InstallModuleAsync(
        InstallModuleRequest request,
        Action<EventEnvelope>? eventHandler = null,
        CancellationToken cancellationToken = default);
}

