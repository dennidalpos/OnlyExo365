using ExchangeAdmin.Contracts;
using ExchangeAdmin.Contracts.Dtos;

namespace ExchangeAdmin.Worker.PowerShell;

public class ExoCommands
{
    private readonly ExoDashboardCommands _dashboardCommands;
    private readonly ExoContactCommands _contactCommands;
    private readonly ExoResourceCommands _resourceCommands;
    private readonly ExoPublicFolderCommands _publicFolderCommands;
    private readonly ExoMobileDeviceCommands _mobileDeviceCommands;
    private readonly ExoMigrationCommands _migrationCommands;
    private readonly ExoSupportCommands _supportCommands;
    private readonly ExoMailboxCommands _mailboxCommands;
    private readonly ExoMailboxReportingCommands _mailboxReportingCommands;
    private readonly ExoMailboxLicenseCommands _mailboxLicenseCommands;
    private readonly ExoMailSecurityCommands _mailSecurityCommands;
    private readonly ExoMailFlowCommands _mailFlowCommands;
    private readonly ExoComplianceCommands _complianceCommands;

    public ExoCommands(PowerShellEngine engine, CapabilityDetector capabilityDetector)
    {
        _mailboxReportingCommands = new ExoMailboxReportingCommands(engine);
        _mailboxLicenseCommands = new ExoMailboxLicenseCommands(engine);
        _mailboxCommands = new ExoMailboxCommands(engine, capabilityDetector, _mailboxReportingCommands, _mailboxLicenseCommands);
        _mailSecurityCommands = new ExoMailSecurityCommands(engine);
        _mailFlowCommands = new ExoMailFlowCommands(engine);
        _complianceCommands = new ExoComplianceCommands(engine);
        _dashboardCommands = new ExoDashboardCommands(engine, capabilityDetector, _mailboxLicenseCommands);
        _contactCommands = new ExoContactCommands(engine);
        _resourceCommands = new ExoResourceCommands(engine, _mailboxReportingCommands);
        _publicFolderCommands = new ExoPublicFolderCommands(engine);
        _mobileDeviceCommands = new ExoMobileDeviceCommands(engine, capabilityDetector);
        _migrationCommands = new ExoMigrationCommands(engine);
        _supportCommands = new ExoSupportCommands(engine);
    }

    public Task<DashboardStatsDto> GetDashboardStatsAsync(
        GetDashboardStatsRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _dashboardCommands.GetDashboardStatsAsync(request, onLog, cancellationToken);

    public Task<GetContactsResponse> GetContactsAsync(
        GetContactsRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _contactCommands.GetContactsAsync(request, onLog, cancellationToken);

    public Task<ContactDetailsDto> GetContactDetailsAsync(
        GetContactDetailsRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _contactCommands.GetContactDetailsAsync(request, onLog, cancellationToken);

    public Task UpsertContactAsync(UpsertContactRequest request, CancellationToken cancellationToken = default)
        => _contactCommands.UpsertContactAsync(request, cancellationToken);

    internal static (string Script, Dictionary<string, object>? Parameters) BuildUpsertContactCommand(UpsertContactRequest request)
        => ExoContactCommands.BuildUpsertContactCommand(request);

    public Task RemoveContactAsync(RemoveContactRequest request, CancellationToken cancellationToken = default)
        => _contactCommands.RemoveContactAsync(request, cancellationToken);

    public Task<GetResourceMailboxesResponse> GetResourceMailboxesAsync(
        GetResourceMailboxesRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _resourceCommands.GetResourceMailboxesAsync(request, onLog, cancellationToken);

    public Task<ResourceMailboxDetailsDto> GetResourceMailboxDetailsAsync(
        GetResourceMailboxDetailsRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _resourceCommands.GetResourceMailboxDetailsAsync(request, onLog, cancellationToken);

    public Task<UpsertResourceMailboxResponse> UpsertResourceMailboxAsync(
        UpsertResourceMailboxRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _resourceCommands.UpsertResourceMailboxAsync(request, onLog, cancellationToken);

    public Task<GetPublicFoldersResponse> GetPublicFoldersAsync(
        GetPublicFoldersRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _publicFolderCommands.GetPublicFoldersAsync(request, onLog, cancellationToken);

    public Task<PublicFolderDetailsDto> GetPublicFolderDetailsAsync(
        GetPublicFolderDetailsRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _publicFolderCommands.GetPublicFolderDetailsAsync(request, onLog, cancellationToken);

    public Task<UpsertPublicFolderResponse> UpsertPublicFolderAsync(
        UpsertPublicFolderRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _publicFolderCommands.UpsertPublicFolderAsync(request, onLog, cancellationToken);

    public Task SetPublicFolderClientPermissionAsync(
        SetPublicFolderClientPermissionRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _publicFolderCommands.SetPublicFolderClientPermissionAsync(request, onLog, cancellationToken);

    public Task RemovePublicFolderAsync(
        RemovePublicFolderRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _publicFolderCommands.RemovePublicFolderAsync(request, onLog, cancellationToken);

    public Task<GetMobileDevicesResponse> GetMobileDevicesAsync(
        GetMobileDevicesRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _mobileDeviceCommands.GetMobileDevicesAsync(request, onLog, cancellationToken);

    public Task<GetMobileDeviceDetailsResponse> GetMobileDeviceDetailsAsync(
        GetMobileDeviceDetailsRequest request,
        Action<string, string>? onLog = null,
        Action<int, int, string>? onProgress = null,
        CancellationToken cancellationToken = default)
        => _mobileDeviceCommands.GetMobileDeviceDetailsAsync(request, onLog, onProgress, cancellationToken);

    internal static string BuildGetMobileDevicesScript(
        int skip,
        int pageSize,
        string escapedSearch,
        string escapedAccessState,
        string sortProperty,
        string sortDirection)
        => ExoMobileDeviceCommands.BuildGetMobileDevicesScript(
            skip,
            pageSize,
            escapedSearch,
            escapedAccessState,
            sortProperty,
            sortDirection);

    internal static int CalculateMobileDevicePageWindowSize(int skip, int pageSize)
        => ExoMobileDeviceCommands.CalculateMobileDevicePageWindowSize(skip, pageSize);

    public Task<GetMobileDeviceMailboxPoliciesResponse> GetMobileDeviceMailboxPoliciesAsync(
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _mobileDeviceCommands.GetMobileDeviceMailboxPoliciesAsync(onLog, cancellationToken);

    internal static string BuildGetMobileDeviceMailboxPoliciesScript()
        => ExoMobileDeviceCommands.BuildGetMobileDeviceMailboxPoliciesScript();

    public Task SetMobileDeviceAccessStateAsync(
        SetMobileDeviceAccessStateRequest request,
        CancellationToken cancellationToken = default)
        => _mobileDeviceCommands.SetMobileDeviceAccessStateAsync(request, cancellationToken);

    public Task ClearMobileDeviceAsync(ClearMobileDeviceRequest request, CancellationToken cancellationToken = default)
        => _mobileDeviceCommands.ClearMobileDeviceAsync(request, cancellationToken);

    public Task SetMobileDeviceMailboxPolicyAsync(SetMobileDeviceMailboxPolicyRequest request, CancellationToken cancellationToken = default)
        => _mobileDeviceCommands.SetMobileDeviceMailboxPolicyAsync(request, cancellationToken);

    public Task<GetMigrationBatchesResponse> GetMigrationBatchesAsync(
        GetMigrationBatchesRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _migrationCommands.GetMigrationBatchesAsync(request, onLog, cancellationToken);

    public Task<GetMigrationEndpointsResponse> GetMigrationEndpointsAsync(
        GetMigrationEndpointsRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _migrationCommands.GetMigrationEndpointsAsync(request, onLog, cancellationToken);

    public Task<MigrationBatchDetailsDto> GetMigrationBatchDetailsAsync(
        GetMigrationBatchDetailsRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _migrationCommands.GetMigrationBatchDetailsAsync(request, onLog, cancellationToken);

    public Task UpsertMigrationEndpointAsync(
        UpsertMigrationEndpointRequest request,
        CancellationToken cancellationToken = default)
        => _migrationCommands.UpsertMigrationEndpointAsync(request, cancellationToken);

    public Task<TestMigrationEndpointResponse> TestMigrationEndpointAsync(
        TestMigrationEndpointRequest request,
        CancellationToken cancellationToken = default)
        => _migrationCommands.TestMigrationEndpointAsync(request, cancellationToken);

    public Task<GetMigrationBatchPreflightResponse> GetMigrationBatchPreflightAsync(
        GetMigrationBatchPreflightRequest request,
        CancellationToken cancellationToken = default)
        => _migrationCommands.GetMigrationBatchPreflightAsync(request, cancellationToken);

    public Task CreateMigrationBatchAsync(
        CreateMigrationBatchRequest request,
        CancellationToken cancellationToken = default)
        => _migrationCommands.CreateMigrationBatchAsync(request, cancellationToken);

    public Task StartMigrationBatchAsync(StartMigrationBatchRequest request, CancellationToken cancellationToken = default)
        => _migrationCommands.StartMigrationBatchAsync(request, cancellationToken);

    public Task CompleteMigrationBatchAsync(CompleteMigrationBatchRequest request, CancellationToken cancellationToken = default)
        => _migrationCommands.CompleteMigrationBatchAsync(request, cancellationToken);

    public Task RemoveMigrationBatchAsync(RemoveMigrationBatchRequest request, CancellationToken cancellationToken = default)
        => _migrationCommands.RemoveMigrationBatchAsync(request, cancellationToken);

    public Task<GetMailboxesResponse> GetMailboxesAsync(
        GetMailboxesRequest request,
        Action<string, string>? onLog = null,
        Action<MailboxListItemDto>? onPartialOutput = null,
        CancellationToken cancellationToken = default)
        => _mailboxCommands.GetMailboxesAsync(request, onLog, onPartialOutput, cancellationToken);

    public Task<GetMailboxProvisioningCandidatesResponse> GetMailboxProvisioningCandidatesAsync(
        GetMailboxProvisioningCandidatesRequest request,
        CancellationToken cancellationToken = default)
        => _mailboxCommands.GetMailboxProvisioningCandidatesAsync(request, cancellationToken);

    public Task<GetDeletedMailboxesResponse> GetDeletedMailboxesAsync(
        GetDeletedMailboxesRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _mailboxCommands.GetDeletedMailboxesAsync(request, onLog, cancellationToken);

    public Task<MailboxDetailsDto> GetMailboxDetailsAsync(
        GetMailboxDetailsRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _mailboxCommands.GetMailboxDetailsAsync(request, onLog, cancellationToken);

    public Task<List<RetentionPolicySummaryDto>> GetRetentionPoliciesAsync(
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _mailboxCommands.GetRetentionPoliciesAsync(onLog, cancellationToken);

    public Task SetMailboxSettingsAsync(
        UpdateMailboxSettingsRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
        => _mailboxCommands.SetMailboxSettingsAsync(request, onLog, cancellationToken);

    public Task SetRetentionPolicyAsync(
        SetRetentionPolicyRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
        => _mailboxCommands.SetRetentionPolicyAsync(request, onLog, cancellationToken);

    public Task SetMailboxAutoReplyConfigurationAsync(
        SetMailboxAutoReplyConfigurationRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
        => _mailboxCommands.SetMailboxAutoReplyConfigurationAsync(request, onLog, cancellationToken);

    public Task CreateMailboxAsync(
        CreateMailboxRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
        => _mailboxCommands.CreateMailboxAsync(request, onLog, cancellationToken);

    public Task ConvertMailboxToSharedAsync(
        ConvertMailboxToSharedRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
        => _mailboxCommands.ConvertMailboxToSharedAsync(request, onLog, cancellationToken);

    public Task ConvertMailboxToRegularAsync(
        ConvertMailboxToRegularRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
        => _mailboxCommands.ConvertMailboxToRegularAsync(request, onLog, cancellationToken);

    public Task<RestoreMailboxResponse> RestoreMailboxAsync(
        RestoreMailboxRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
        => _mailboxCommands.RestoreMailboxAsync(request, onLog, cancellationToken);

    public Task<GetMailboxSpaceReportResponse> GetMailboxSpaceReportAsync(
        GetMailboxSpaceReportRequest request,
        Action<string, string>? onLog,
        Action<int, int>? onProgress,
        CancellationToken cancellationToken)
        => _mailboxReportingCommands.GetMailboxSpaceReportAsync(request, onLog, onProgress, cancellationToken);

    public Task<GetMailboxAccessReportResponse> GetMailboxAccessReportAsync(
        GetMailboxAccessReportRequest request,
        Action<string, string>? onLog,
        Action<int, int>? onProgress,
        CancellationToken cancellationToken)
        => _mailboxReportingCommands.GetMailboxAccessReportAsync(request, onLog, onProgress, cancellationToken);

    public Task<MailboxPermissionsDto> GetMailboxPermissionsAsync(
        string identity,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
        => _mailboxReportingCommands.GetMailboxPermissionsAsync(identity, onLog, cancellationToken);

    public Task<GetMailboxFolderPermissionsResponse> GetMailboxFolderPermissionsAsync(
        GetMailboxFolderPermissionsRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
        => _mailboxReportingCommands.GetMailboxFolderPermissionsAsync(request, onLog, cancellationToken);

    public Task SetMailboxPermissionAsync(
        SetMailboxPermissionRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
        => _mailboxReportingCommands.SetMailboxPermissionAsync(request, onLog, cancellationToken);

    public Task SetMailboxFolderPermissionAsync(
        SetMailboxFolderPermissionRequest request,
        Action<string, string>? onLog,
        CancellationToken cancellationToken)
        => _mailboxReportingCommands.SetMailboxFolderPermissionAsync(request, onLog, cancellationToken);

    public Task<ApplyPermissionsDeltaPlanResponse> ApplyPermissionsDeltaPlanAsync(
        ApplyPermissionsDeltaPlanRequest request,
        Action<string, string>? onLog,
        Action<int, int>? onProgress,
        CancellationToken cancellationToken)
        => _mailboxReportingCommands.ApplyPermissionsDeltaPlanAsync(request, onLog, onProgress, cancellationToken);

    public Task<List<TenantLicenseDto>> GetTenantLicensesAsync(CancellationToken cancellationToken = default)
        => _mailboxLicenseCommands.GetTenantLicensesAsync(cancellationToken);

    public Task<List<AdminRoleMemberDto>> GetAdminRoleMembersAsync(CancellationToken cancellationToken = default)
        => _mailboxLicenseCommands.GetAdminRoleMembersAsync(cancellationToken);

    public Task<GetMessageTraceResponse> GetMessageTraceAsync(GetMessageTraceRequest request, CancellationToken cancellationToken = default)
        => _mailFlowCommands.GetMessageTraceAsync(request, cancellationToken);

    public Task<GetMessageTraceDetailsResponse> GetMessageTraceDetailsAsync(GetMessageTraceDetailsRequest request, CancellationToken cancellationToken = default)
        => _mailFlowCommands.GetMessageTraceDetailsAsync(request, cancellationToken);

    public Task<GetMailSecurityBaselineResponse> GetMailSecurityBaselineAsync(
        GetMailSecurityBaselineRequest request,
        CancellationToken cancellationToken = default)
        => _mailSecurityCommands.GetMailSecurityBaselineAsync(request, cancellationToken);

    public Task UpdateDkimSigningConfigAsync(UpdateDkimSigningConfigRequest request, CancellationToken cancellationToken = default)
        => _mailSecurityCommands.UpdateDkimSigningConfigAsync(request, cancellationToken);

    public Task UpdateHostedContentFilterPolicyAsync(UpdateHostedContentFilterPolicyRequest request, CancellationToken cancellationToken = default)
        => _mailSecurityCommands.UpdateHostedContentFilterPolicyAsync(request, cancellationToken);

    public Task UpdateAntiPhishPolicyAsync(UpdateAntiPhishPolicyRequest request, CancellationToken cancellationToken = default)
        => _mailSecurityCommands.UpdateAntiPhishPolicyAsync(request, cancellationToken);

    public Task UpdateMalwareFilterPolicyAsync(UpdateMalwareFilterPolicyRequest request, CancellationToken cancellationToken = default)
        => _mailSecurityCommands.UpdateMalwareFilterPolicyAsync(request, cancellationToken);

    public Task UpdateQuarantinePolicyAsync(UpdateQuarantinePolicyRequest request, CancellationToken cancellationToken = default)
        => _mailSecurityCommands.UpdateQuarantinePolicyAsync(request, cancellationToken);

    public Task UpdateHostedOutboundSpamFilterPolicyAsync(UpdateHostedOutboundSpamFilterPolicyRequest request, CancellationToken cancellationToken = default)
        => _mailSecurityCommands.UpdateHostedOutboundSpamFilterPolicyAsync(request, cancellationToken);

    public Task<GetComplianceWorkspaceResponse> GetComplianceWorkspaceAsync(
        GetComplianceWorkspaceRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _complianceCommands.GetComplianceWorkspaceAsync(request, onLog, cancellationToken);

    public Task<SearchUnifiedAuditLogResponse> SearchUnifiedAuditLogAsync(
        SearchUnifiedAuditLogRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _complianceCommands.SearchUnifiedAuditLogAsync(request, onLog, cancellationToken);

    public Task CreateComplianceSearchAsync(
        CreateComplianceSearchRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _complianceCommands.CreateComplianceSearchAsync(request, onLog, cancellationToken);

    public Task StartComplianceSearchAsync(
        StartComplianceSearchRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _complianceCommands.StartComplianceSearchAsync(request, onLog, cancellationToken);

    public Task RemoveComplianceSearchAsync(
        RemoveComplianceSearchRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _complianceCommands.RemoveComplianceSearchAsync(request, onLog, cancellationToken);

    public Task<InvokeComplianceActionResponse> InvokeComplianceActionAsync(
        InvokeComplianceActionRequest request,
        Action<string, string>? onLog = null,
        CancellationToken cancellationToken = default)
        => _complianceCommands.InvokeComplianceActionAsync(request, onLog, cancellationToken);

    internal static string BuildConnectComplianceCommand(ExchangeOnlineConfiguration configuration)
        => ComplianceCommandBuilder.BuildConnectComplianceCommand(configuration);

    internal static string BuildConnectComplianceSearchOnlyCommand(ExchangeOnlineConfiguration configuration)
        => ComplianceCommandBuilder.BuildConnectComplianceSearchOnlyCommand(configuration);

    internal static string BuildCreateComplianceSearchScript(CreateComplianceSearchRequest request)
        => ComplianceCommandBuilder.BuildCreateComplianceSearchScript(request);

    internal static string BuildInvokeComplianceActionScript(
        InvokeComplianceActionRequest request,
        IReadOnlyList<string> exchangeLocations,
        string? contentMatchQuery)
        => ComplianceCommandBuilder.BuildInvokeComplianceActionScript(request, exchangeLocations, contentMatchQuery);

    internal static string BuildUpdateHostedContentFilterPolicyScript(UpdateHostedContentFilterPolicyRequest request)
        => MailSecurityCommandBuilder.BuildUpdateHostedContentFilterPolicyScript(request);

    internal static string BuildUpdateAntiPhishPolicyScript(UpdateAntiPhishPolicyRequest request)
        => MailSecurityCommandBuilder.BuildUpdateAntiPhishPolicyScript(request);

    internal static string BuildUpdateHostedOutboundSpamFilterPolicyScript(UpdateHostedOutboundSpamFilterPolicyRequest request)
        => MailSecurityCommandBuilder.BuildUpdateHostedOutboundSpamFilterPolicyScript(request);

    internal static string BuildUpsertPublicFolderScript(UpsertPublicFolderRequest request)
        => ExoPublicFolderCommands.BuildUpsertPublicFolderScript(request);

    internal static string BuildSetPublicFolderClientPermissionScript(SetPublicFolderClientPermissionRequest request)
        => ExoPublicFolderCommands.BuildSetPublicFolderClientPermissionScript(request);

    internal static string BuildRemovePublicFolderScript(RemovePublicFolderRequest request)
        => ExoPublicFolderCommands.BuildRemovePublicFolderScript(request);

    internal static string BuildGetMailboxFolderPermissionsScript(GetMailboxFolderPermissionsRequest request)
        => ExoMailboxReportingCommands.BuildGetMailboxFolderPermissionsScript(request);

    internal static string BuildSetMailboxFolderPermissionScript(SetMailboxFolderPermissionRequest request)
        => ExoMailboxReportingCommands.BuildSetMailboxFolderPermissionScript(request);

    internal static (string Script, Dictionary<string, object>? Parameters) BuildUpsertMigrationEndpointCommand(UpsertMigrationEndpointRequest request)
        => MigrationCommandBuilder.BuildUpsertMigrationEndpointCommand(request);

    internal static (string Script, Dictionary<string, object>? Parameters) BuildTestMigrationEndpointCommand(TestMigrationEndpointRequest request)
        => MigrationCommandBuilder.BuildTestMigrationEndpointCommand(request);

    internal static (string Script, Dictionary<string, object>? Parameters) BuildCreateMigrationBatchCommand(CreateMigrationBatchRequest request)
        => MigrationCommandBuilder.BuildCreateMigrationBatchCommand(request);

    public Task<GetTransportRulesResponse> GetTransportRulesAsync(CancellationToken cancellationToken = default)
        => _mailFlowCommands.GetTransportRulesAsync(cancellationToken);

    public Task SetTransportRuleStateAsync(SetTransportRuleStateRequest request, CancellationToken cancellationToken = default)
        => _mailFlowCommands.SetTransportRuleStateAsync(request, cancellationToken);

    public Task<GetConnectorsResponse> GetConnectorsAsync(CancellationToken cancellationToken = default)
        => _mailFlowCommands.GetConnectorsAsync(cancellationToken);

    public Task<GetAcceptedDomainsResponse> GetAcceptedDomainsAsync(CancellationToken cancellationToken = default)
        => _mailFlowCommands.GetAcceptedDomainsAsync(cancellationToken);

    public Task UpsertTransportRuleAsync(UpsertTransportRuleRequest request, CancellationToken cancellationToken = default)
        => _mailFlowCommands.UpsertTransportRuleAsync(request, cancellationToken);

    public Task RemoveTransportRuleAsync(RemoveTransportRuleRequest request, CancellationToken cancellationToken = default)
        => _mailFlowCommands.RemoveTransportRuleAsync(request, cancellationToken);

    public Task<TestTransportRuleResponse> TestTransportRuleAsync(TestTransportRuleRequest request, CancellationToken cancellationToken = default)
        => _mailFlowCommands.TestTransportRuleAsync(request, cancellationToken);

    public Task UpsertConnectorAsync(UpsertConnectorRequest request, CancellationToken cancellationToken = default)
        => _mailFlowCommands.UpsertConnectorAsync(request, cancellationToken);

    public Task RemoveConnectorAsync(RemoveConnectorRequest request, CancellationToken cancellationToken = default)
        => _mailFlowCommands.RemoveConnectorAsync(request, cancellationToken);

    public Task UpsertAcceptedDomainAsync(UpsertAcceptedDomainRequest request, CancellationToken cancellationToken = default)
        => _mailFlowCommands.UpsertAcceptedDomainAsync(request, cancellationToken);

    public Task RemoveAcceptedDomainAsync(RemoveAcceptedDomainRequest request, CancellationToken cancellationToken = default)
        => _mailFlowCommands.RemoveAcceptedDomainAsync(request, cancellationToken);

    public Task<GetRemoteDomainsResponse> GetRemoteDomainsAsync(CancellationToken cancellationToken = default)
        => _mailFlowCommands.GetRemoteDomainsAsync(cancellationToken);

    public Task UpsertRemoteDomainAsync(UpsertRemoteDomainRequest request, CancellationToken cancellationToken = default)
        => _mailFlowCommands.UpsertRemoteDomainAsync(request, cancellationToken);

    public Task RemoveRemoteDomainAsync(RemoveRemoteDomainRequest request, CancellationToken cancellationToken = default)
        => _mailFlowCommands.RemoveRemoteDomainAsync(request, cancellationToken);

    public Task<GetOrganizationRelationshipsResponse> GetOrganizationRelationshipsAsync(CancellationToken cancellationToken = default)
        => _mailFlowCommands.GetOrganizationRelationshipsAsync(cancellationToken);

    public Task UpsertOrganizationRelationshipAsync(UpsertOrganizationRelationshipRequest request, CancellationToken cancellationToken = default)
        => _mailFlowCommands.UpsertOrganizationRelationshipAsync(request, cancellationToken);

    public Task RemoveOrganizationRelationshipAsync(RemoveOrganizationRelationshipRequest request, CancellationToken cancellationToken = default)
        => _mailFlowCommands.RemoveOrganizationRelationshipAsync(request, cancellationToken);

    public Task<GetAddressListsResponse> GetAddressListsAsync(CancellationToken cancellationToken = default)
        => _mailFlowCommands.GetAddressListsAsync(cancellationToken);

    public Task UpsertAddressListAsync(UpsertAddressListRequest request, CancellationToken cancellationToken = default)
        => _mailFlowCommands.UpsertAddressListAsync(request, cancellationToken);

    public Task RemoveAddressListAsync(RemoveAddressListRequest request, CancellationToken cancellationToken = default)
        => _mailFlowCommands.RemoveAddressListAsync(request, cancellationToken);

    public Task<GetAddressBookPoliciesResponse> GetAddressBookPoliciesAsync(CancellationToken cancellationToken = default)
        => _mailFlowCommands.GetAddressBookPoliciesAsync(cancellationToken);

    public Task UpsertAddressBookPolicyAsync(UpsertAddressBookPolicyRequest request, CancellationToken cancellationToken = default)
        => _mailFlowCommands.UpsertAddressBookPolicyAsync(request, cancellationToken);

    public Task RemoveAddressBookPolicyAsync(RemoveAddressBookPolicyRequest request, CancellationToken cancellationToken = default)
        => _mailFlowCommands.RemoveAddressBookPolicyAsync(request, cancellationToken);

    public Task<GetOfflineAddressBooksResponse> GetOfflineAddressBooksAsync(CancellationToken cancellationToken = default)
        => _mailFlowCommands.GetOfflineAddressBooksAsync(cancellationToken);

    public Task UpsertOfflineAddressBookAsync(UpsertOfflineAddressBookRequest request, CancellationToken cancellationToken = default)
        => _mailFlowCommands.UpsertOfflineAddressBookAsync(request, cancellationToken);

    public Task RemoveOfflineAddressBookAsync(RemoveOfflineAddressBookRequest request, CancellationToken cancellationToken = default)
        => _mailFlowCommands.RemoveOfflineAddressBookAsync(request, cancellationToken);

    public Task<GetSharingPoliciesResponse> GetSharingPoliciesAsync(CancellationToken cancellationToken = default)
        => _mailFlowCommands.GetSharingPoliciesAsync(cancellationToken);

    public Task UpsertSharingPolicyAsync(UpsertSharingPolicyRequest request, CancellationToken cancellationToken = default)
        => _mailFlowCommands.UpsertSharingPolicyAsync(request, cancellationToken);

    public Task RemoveSharingPolicyAsync(RemoveSharingPolicyRequest request, CancellationToken cancellationToken = default)
        => _mailFlowCommands.RemoveSharingPolicyAsync(request, cancellationToken);

    public Task<GetUserLicensesResponse> GetUserLicensesAsync(string userPrincipalName, CancellationToken cancellationToken = default)
        => _mailboxLicenseCommands.GetUserLicensesAsync(userPrincipalName, cancellationToken);

    public Task SetUserLicenseAsync(SetUserLicenseRequest request, CancellationToken cancellationToken = default)
        => _mailboxLicenseCommands.SetUserLicenseAsync(request, cancellationToken);

    public Task<GetUsageLocationSuggestionResponse> GetUsageLocationSuggestionAsync(
        GetUsageLocationSuggestionRequest request,
        CancellationToken cancellationToken = default)
        => _mailboxLicenseCommands.GetUsageLocationSuggestionAsync(request, cancellationToken);

    public Task SetUserUsageLocationAsync(
        SetUserUsageLocationRequest request,
        CancellationToken cancellationToken = default)
        => _mailboxLicenseCommands.SetUserUsageLocationAsync(request, cancellationToken);

    public Task<GetAvailableLicensesResponse> GetAvailableLicensesAsync(CancellationToken cancellationToken = default)
        => _mailboxLicenseCommands.GetAvailableLicensesAsync(cancellationToken);

    public Task<PrerequisiteStatusDto> CheckPrerequisitesAsync(CancellationToken cancellationToken = default)
        => _supportCommands.CheckPrerequisitesAsync(cancellationToken);

    public Task<InstallModuleResponse> InstallModuleAsync(InstallModuleRequest request, CancellationToken cancellationToken = default)
        => _supportCommands.InstallModuleAsync(request, cancellationToken);
}
