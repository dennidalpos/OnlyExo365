using ExchangeAdmin.Application.Services;
using ExchangeAdmin.Contracts.Dtos;
using ExchangeAdmin.Contracts.Messages;
using ExchangeAdmin.Domain.Results;
using ExchangeAdmin.Infrastructure.Ipc;

namespace ExchangeAdmin.Tests;

public abstract class TestWorkerServiceBase : IWorkerService
{
    private event EventHandler<WorkerConnectionState>? StateChangedHandlers;
    private event EventHandler<EventEnvelope>? EventReceivedHandlers;
    private event EventHandler<CapabilityMapDto>? CapabilitiesUpdatedHandlers;

    public virtual WorkerConnectionState ConnectionState => WorkerConnectionState.Connected;

    public virtual WorkerStatus Status { get; } = new()
    {
        State = WorkerConnectionState.Connected,
        IsModuleAvailable = true
    };

    public virtual CapabilityMapDto? Capabilities => null;

    public virtual event EventHandler<WorkerConnectionState>? StateChanged
    {
        add => StateChangedHandlers += value;
        remove => StateChangedHandlers -= value;
    }

    public virtual event EventHandler<EventEnvelope>? EventReceived
    {
        add => EventReceivedHandlers += value;
        remove => EventReceivedHandlers -= value;
    }

    public virtual event EventHandler<CapabilityMapDto>? CapabilitiesUpdated
    {
        add => CapabilitiesUpdatedHandlers += value;
        remove => CapabilitiesUpdatedHandlers -= value;
    }

    protected void PublishStateChanged(WorkerConnectionState state)
        => StateChangedHandlers?.Invoke(this, state);

    protected void PublishEventReceived(EventEnvelope envelope)
        => EventReceivedHandlers?.Invoke(this, envelope);

    protected void PublishCapabilitiesUpdated(CapabilityMapDto capabilities)
        => CapabilitiesUpdatedHandlers?.Invoke(this, capabilities);

    public virtual Task<bool> StartWorkerAsync(CancellationToken cancellationToken = default) => Task.FromResult(true);
    public virtual Task StopWorkerAsync() => Task.CompletedTask;
    public virtual Task<bool> RestartWorkerAsync(CancellationToken cancellationToken = default) => Task.FromResult(true);
    public virtual void KillWorker() { }
    public virtual Task<Result<SetWorkerConsoleVisibilityResponse>> SetWorkerConsoleVisibilityAsync(bool isVisible, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<ConnectionStatusDto>> ConnectExchangeAsync(Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> DisconnectExchangeAsync(CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<ConnectionStatusDto>> GetConnectionStatusAsync(CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<CapabilityMapDto>> DetectCapabilitiesAsync(bool forceRefresh = false, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();

    public virtual Task<Result<DashboardStatsDto>> GetDashboardStatsAsync(GetDashboardStatsRequest? request = null, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetContactsResponse>> GetContactsAsync(GetContactsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<ContactDetailsDto>> GetContactDetailsAsync(GetContactDetailsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpsertContactAsync(UpsertContactRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> RemoveContactAsync(RemoveContactRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();

    public virtual Task<Result<GetResourceMailboxesResponse>> GetResourceMailboxesAsync(GetResourceMailboxesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<ResourceMailboxDetailsDto>> GetResourceMailboxDetailsAsync(GetResourceMailboxDetailsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<UpsertResourceMailboxResponse>> UpsertResourceMailboxAsync(UpsertResourceMailboxRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<ApplyPermissionsDeltaPlanResponse>> ApplyPermissionsDeltaPlanAsync(ApplyPermissionsDeltaPlanRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();

    public virtual Task<Result<GetMigrationBatchesResponse>> GetMigrationBatchesAsync(GetMigrationBatchesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetMigrationEndpointsResponse>> GetMigrationEndpointsAsync(GetMigrationEndpointsRequest? request = null, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<MigrationBatchDetailsDto>> GetMigrationBatchDetailsAsync(GetMigrationBatchDetailsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpsertMigrationEndpointAsync(UpsertMigrationEndpointRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<TestMigrationEndpointResponse>> TestMigrationEndpointAsync(TestMigrationEndpointRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetMigrationBatchPreflightResponse>> GetMigrationBatchPreflightAsync(GetMigrationBatchPreflightRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> CreateMigrationBatchAsync(CreateMigrationBatchRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> StartMigrationBatchAsync(StartMigrationBatchRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> CompleteMigrationBatchAsync(CompleteMigrationBatchRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> RemoveMigrationBatchAsync(RemoveMigrationBatchRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();

    public virtual Task<Result<GetMailboxProvisioningCandidatesResponse>> GetMailboxProvisioningCandidatesAsync(GetMailboxProvisioningCandidatesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetMailboxesResponse>> GetMailboxesAsync(GetMailboxesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetDeletedMailboxesResponse>> GetDeletedMailboxesAsync(GetDeletedMailboxesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<MailboxDetailsDto>> GetMailboxDetailsAsync(GetMailboxDetailsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetRetentionPoliciesResponse>> GetRetentionPoliciesAsync(GetRetentionPoliciesRequest? request = null, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> SetRetentionPolicyAsync(SetRetentionPolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetMailboxFolderPermissionsResponse>> GetMailboxFolderPermissionsAsync(GetMailboxFolderPermissionsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> SetMailboxFolderPermissionAsync(SetMailboxFolderPermissionRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpdateMailboxSettingsAsync(UpdateMailboxSettingsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> SetMailboxAutoReplyConfigurationAsync(SetMailboxAutoReplyConfigurationRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> ConvertMailboxToSharedAsync(ConvertMailboxToSharedRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> ConvertMailboxToRegularAsync(ConvertMailboxToRegularRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<RestoreMailboxResponse>> RestoreMailboxAsync(RestoreMailboxRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetMailboxSpaceReportResponse>> GetMailboxSpaceReportAsync(GetMailboxSpaceReportRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetMailboxAccessReportResponse>> GetMailboxAccessReportAsync(GetMailboxAccessReportRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> CreateMailboxAsync(CreateMailboxRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetAcceptedDomainsResponse>> GetAcceptedDomainsAsync(GetAcceptedDomainsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetUserLicensesResponse>> GetUserLicensesAsync(GetUserLicensesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> SetUserLicenseAsync(SetUserLicenseRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetUsageLocationSuggestionResponse>> GetUsageLocationSuggestionAsync(GetUsageLocationSuggestionRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> SetUserUsageLocationAsync(SetUserUsageLocationRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetAvailableLicensesResponse>> GetAvailableLicensesAsync(Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();

    public virtual Task<Result<GetDistributionListsResponse>> GetDistributionListsAsync(GetDistributionListsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<DistributionListDetailsDto>> GetDistributionListDetailsAsync(GetDistributionListDetailsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GroupMembersPageDto>> GetGroupMembersAsync(GetGroupMembersRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> ModifyGroupMemberAsync(ModifyGroupMemberRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> SetDistributionListSettingsAsync(SetDistributionListSettingsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> CreateDistributionListAsync(CreateDistributionListRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<PreviewDynamicGroupMembersResponse>> PreviewDynamicGroupMembersAsync(PreviewDynamicGroupMembersRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();

    public virtual Task<Result<GetMailSecurityBaselineResponse>> GetMailSecurityBaselineAsync(GetMailSecurityBaselineRequest? request = null, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpdateDkimSigningConfigAsync(UpdateDkimSigningConfigRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpdateHostedContentFilterPolicyAsync(UpdateHostedContentFilterPolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpdateAntiPhishPolicyAsync(UpdateAntiPhishPolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpdateMalwareFilterPolicyAsync(UpdateMalwareFilterPolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpdateQuarantinePolicyAsync(UpdateQuarantinePolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpdateHostedOutboundSpamFilterPolicyAsync(UpdateHostedOutboundSpamFilterPolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();

    public virtual Task<Result<GetTransportRulesResponse>> GetTransportRulesAsync(GetTransportRulesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> SetTransportRuleStateAsync(SetTransportRuleStateRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpsertTransportRuleAsync(UpsertTransportRuleRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> RemoveTransportRuleAsync(RemoveTransportRuleRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<TestTransportRuleResponse>> TestTransportRuleAsync(TestTransportRuleRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetConnectorsResponse>> GetConnectorsAsync(GetConnectorsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpsertConnectorAsync(UpsertConnectorRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> RemoveConnectorAsync(RemoveConnectorRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpsertAcceptedDomainAsync(UpsertAcceptedDomainRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> RemoveAcceptedDomainAsync(RemoveAcceptedDomainRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetRemoteDomainsResponse>> GetRemoteDomainsAsync(GetRemoteDomainsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpsertRemoteDomainAsync(UpsertRemoteDomainRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> RemoveRemoteDomainAsync(RemoveRemoteDomainRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetOrganizationRelationshipsResponse>> GetOrganizationRelationshipsAsync(GetOrganizationRelationshipsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetAddressListsResponse>> GetAddressListsAsync(GetAddressListsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpsertAddressListAsync(UpsertAddressListRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> RemoveAddressListAsync(RemoveAddressListRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetAddressBookPoliciesResponse>> GetAddressBookPoliciesAsync(GetAddressBookPoliciesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpsertAddressBookPolicyAsync(UpsertAddressBookPolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> RemoveAddressBookPolicyAsync(RemoveAddressBookPolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetOfflineAddressBooksResponse>> GetOfflineAddressBooksAsync(GetOfflineAddressBooksRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpsertOfflineAddressBookAsync(UpsertOfflineAddressBookRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> RemoveOfflineAddressBookAsync(RemoveOfflineAddressBookRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetSharingPoliciesResponse>> GetSharingPoliciesAsync(GetSharingPoliciesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpsertSharingPolicyAsync(UpsertSharingPolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> RemoveSharingPolicyAsync(RemoveSharingPolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpsertOrganizationRelationshipAsync(UpsertOrganizationRelationshipRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> RemoveOrganizationRelationshipAsync(RemoveOrganizationRelationshipRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();

    public virtual Task<Result<GetComplianceWorkspaceResponse>> GetComplianceWorkspaceAsync(GetComplianceWorkspaceRequest? request = null, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<SearchUnifiedAuditLogResponse>> SearchUnifiedAuditLogAsync(SearchUnifiedAuditLogRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> CreateComplianceSearchAsync(CreateComplianceSearchRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> StartComplianceSearchAsync(StartComplianceSearchRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> RemoveComplianceSearchAsync(RemoveComplianceSearchRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<InvokeComplianceActionResponse>> InvokeComplianceActionAsync(InvokeComplianceActionRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();

    public virtual Task<Result<GetPublicFoldersResponse>> GetPublicFoldersAsync(GetPublicFoldersRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<PublicFolderDetailsDto>> GetPublicFolderDetailsAsync(GetPublicFolderDetailsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<UpsertPublicFolderResponse>> UpsertPublicFolderAsync(UpsertPublicFolderRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> SetPublicFolderClientPermissionAsync(SetPublicFolderClientPermissionRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> RemovePublicFolderAsync(RemovePublicFolderRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();

    public virtual Task<Result<GetMobileDevicesResponse>> GetMobileDevicesAsync(GetMobileDevicesRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetMobileDeviceDetailsResponse>> GetMobileDeviceDetailsAsync(GetMobileDeviceDetailsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetMobileDeviceMailboxPoliciesResponse>> GetMobileDeviceMailboxPoliciesAsync(GetMobileDeviceMailboxPoliciesRequest? request = null, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> SetMobileDeviceAccessStateAsync(SetMobileDeviceAccessStateRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> ClearMobileDeviceAsync(ClearMobileDeviceRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> SetMobileDeviceMailboxPolicyAsync(SetMobileDeviceMailboxPolicyRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();

    public virtual Task<Result<GetRoleGroupsResponse>> GetRoleGroupsAsync(GetRoleGroupsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<RoleGroupDetailsDto>> GetRoleGroupDetailsAsync(GetRoleGroupDetailsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> UpsertRoleGroupAsync(UpsertRoleGroupRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result> ModifyRoleGroupMemberAsync(ModifyRoleGroupMemberRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();

    public virtual Task<Result<GetMessageTraceResponse>> GetMessageTraceAsync(GetMessageTraceRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<GetMessageTraceDetailsResponse>> GetMessageTraceDetailsAsync(GetMessageTraceDetailsRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();

    public virtual Task<Result<PrerequisiteStatusDto>> CheckPrerequisitesAsync(Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();
    public virtual Task<Result<InstallModuleResponse>> InstallModuleAsync(InstallModuleRequest request, Action<EventEnvelope>? eventHandler = null, CancellationToken cancellationToken = default) => throw TestStubExceptions.CreateUnsupported();

    public virtual Task CancelOperationAsync(string correlationId) => Task.CompletedTask;
}
