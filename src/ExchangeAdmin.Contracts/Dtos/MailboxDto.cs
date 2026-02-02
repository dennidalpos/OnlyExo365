using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Dtos;

public class MailboxListItemDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("guid")]
    public string? Guid { get; set; }

    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;

    [JsonPropertyName("primarySmtpAddress")]
    public string PrimarySmtpAddress { get; set; } = string.Empty;

    [JsonPropertyName("recipientType")]
    public string RecipientType { get; set; } = string.Empty;

    [JsonPropertyName("recipientTypeDetails")]
    public string RecipientTypeDetails { get; set; } = string.Empty;

    [JsonPropertyName("alias")]
    public string? Alias { get; set; }

    [JsonPropertyName("isInactiveMailbox")]
    public bool IsInactiveMailbox { get; set; }
}

public class MailboxDto : MailboxListItemDto
{
    [JsonPropertyName("userPrincipalName")]
    public string? UserPrincipalName { get; set; }

    [JsonPropertyName("samAccountName")]
    public string? SamAccountName { get; set; }

    [JsonPropertyName("organizationalUnit")]
    public string? OrganizationalUnit { get; set; }

    [JsonPropertyName("whenCreated")]
    public DateTime? WhenCreated { get; set; }

    [JsonPropertyName("whenMailboxCreated")]
    public DateTime? WhenMailboxCreated { get; set; }
}

public class MailboxDetailsDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("guid")]
    public string? Guid { get; set; }

    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;

    [JsonPropertyName("primarySmtpAddress")]
    public string PrimarySmtpAddress { get; set; } = string.Empty;

    [JsonPropertyName("userPrincipalName")]
    public string? UserPrincipalName { get; set; }

    [JsonPropertyName("alias")]
    public string? Alias { get; set; }

    [JsonPropertyName("recipientType")]
    public string RecipientType { get; set; } = string.Empty;

    [JsonPropertyName("recipientTypeDetails")]
    public string RecipientTypeDetails { get; set; } = string.Empty;

    [JsonPropertyName("emailAddresses")]
    public List<string> EmailAddresses { get; set; } = new();

    [JsonPropertyName("retentionPolicy")]
    public string? RetentionPolicy { get; set; }

    [JsonPropertyName("features")]
    public MailboxFeaturesDto Features { get; set; } = new();

    [JsonPropertyName("statistics")]
    public MailboxStatisticsDto? Statistics { get; set; }

    [JsonPropertyName("inboxRules")]
    public List<InboxRuleDto>? InboxRules { get; set; }

    [JsonPropertyName("autoReplyConfiguration")]
    public AutoReplyConfigurationDto? AutoReplyConfiguration { get; set; }

    [JsonPropertyName("permissions")]
    public MailboxPermissionsDto? Permissions { get; set; }

    [JsonPropertyName("assignedLicenses")]
    public List<UserLicenseDto> AssignedLicenses { get; set; } = new();

    [JsonPropertyName("whenCreated")]
    public DateTime? WhenCreated { get; set; }

    [JsonPropertyName("whenMailboxCreated")]
    public DateTime? WhenMailboxCreated { get; set; }
}

public class RetentionPolicySummaryDto
{
    [JsonPropertyName("id")]
    public string? Id { get; set; }

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("description")]
    public string? Description { get; set; }

    [JsonPropertyName("requiresArchive")]
    public bool RequiresArchive { get; set; }
}

public class MailboxFeaturesDto
{
    [JsonPropertyName("archiveEnabled")]
    public bool ArchiveEnabled { get; set; }

    [JsonPropertyName("archiveName")]
    public string? ArchiveName { get; set; }

    [JsonPropertyName("archiveGuid")]
    public string? ArchiveGuid { get; set; }

    [JsonPropertyName("archiveStatus")]
    public string? ArchiveStatus { get; set; }

    [JsonPropertyName("litigationHoldEnabled")]
    public bool LitigationHoldEnabled { get; set; }

    [JsonPropertyName("litigationHoldDate")]
    public DateTime? LitigationHoldDate { get; set; }

    [JsonPropertyName("litigationHoldOwner")]
    public string? LitigationHoldOwner { get; set; }

    [JsonPropertyName("litigationHoldDuration")]
    public string? LitigationHoldDuration { get; set; }

    [JsonPropertyName("auditEnabled")]
    public bool AuditEnabled { get; set; }

    [JsonPropertyName("auditLogAgeLimit")]
    public string? AuditLogAgeLimit { get; set; }

    [JsonPropertyName("auditAdmin")]
    public List<string>? AuditAdmin { get; set; }

    [JsonPropertyName("auditDelegate")]
    public List<string>? AuditDelegate { get; set; }

    [JsonPropertyName("auditOwner")]
    public List<string>? AuditOwner { get; set; }

    [JsonPropertyName("forwardingAddress")]
    public string? ForwardingAddress { get; set; }

    [JsonPropertyName("forwardingSmtpAddress")]
    public string? ForwardingSmtpAddress { get; set; }

    [JsonPropertyName("deliverToMailboxAndForward")]
    public bool DeliverToMailboxAndForward { get; set; }

    [JsonPropertyName("prohibitSendQuota")]
    public string? ProhibitSendQuota { get; set; }

    [JsonPropertyName("prohibitSendQuotaBytes")]
    public long? ProhibitSendQuotaBytes { get; set; }

    [JsonPropertyName("prohibitSendReceiveQuota")]
    public string? ProhibitSendReceiveQuota { get; set; }

    [JsonPropertyName("prohibitSendReceiveQuotaBytes")]
    public long? ProhibitSendReceiveQuotaBytes { get; set; }

    [JsonPropertyName("issueWarningQuota")]
    public string? IssueWarningQuota { get; set; }

    [JsonPropertyName("issueWarningQuotaBytes")]
    public long? IssueWarningQuotaBytes { get; set; }

    [JsonPropertyName("maxSendSize")]
    public string? MaxSendSize { get; set; }

    [JsonPropertyName("maxReceiveSize")]
    public string? MaxReceiveSize { get; set; }

    [JsonPropertyName("retentionHoldEnabled")]
    public bool RetentionHoldEnabled { get; set; }

    [JsonPropertyName("singleItemRecoveryEnabled")]
    public bool SingleItemRecoveryEnabled { get; set; }

    [JsonPropertyName("retainDeletedItemsFor")]
    public string? RetainDeletedItemsFor { get; set; }

    [JsonPropertyName("hiddenFromAddressListsEnabled")]
    public bool HiddenFromAddressListsEnabled { get; set; }
}

public class MailboxStatisticsDto
{
    [JsonPropertyName("totalItemSize")]
    public string? TotalItemSize { get; set; }

    [JsonPropertyName("totalItemSizeBytes")]
    public long? TotalItemSizeBytes { get; set; }

    [JsonPropertyName("itemCount")]
    public int ItemCount { get; set; }

    [JsonPropertyName("deletedItemCount")]
    public int DeletedItemCount { get; set; }

    [JsonPropertyName("totalDeletedItemSize")]
    public string? TotalDeletedItemSize { get; set; }

    [JsonPropertyName("lastLogonTime")]
    public DateTime? LastLogonTime { get; set; }

    [JsonPropertyName("lastLogoffTime")]
    public DateTime? LastLogoffTime { get; set; }
}

public class InboxRuleDto
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("ruleIdentity")]
    public string? RuleIdentity { get; set; }

    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; }

    [JsonPropertyName("priority")]
    public int Priority { get; set; }

    [JsonPropertyName("description")]
    public string? Description { get; set; }

    [JsonPropertyName("forwardTo")]
    public List<string>? ForwardTo { get; set; }

    [JsonPropertyName("forwardAsAttachmentTo")]
    public List<string>? ForwardAsAttachmentTo { get; set; }

    [JsonPropertyName("redirectTo")]
    public List<string>? RedirectTo { get; set; }

    [JsonPropertyName("deleteMessage")]
    public bool DeleteMessage { get; set; }

    [JsonPropertyName("moveToFolder")]
    public string? MoveToFolder { get; set; }
}

public class AutoReplyConfigurationDto
{
    [JsonPropertyName("autoReplyState")]
    public string AutoReplyState { get; set; } = "Disabled";

    [JsonPropertyName("startTime")]
    public DateTime? StartTime { get; set; }

    [JsonPropertyName("endTime")]
    public DateTime? EndTime { get; set; }

    [JsonPropertyName("internalMessage")]
    public string? InternalMessage { get; set; }

    [JsonPropertyName("externalMessage")]
    public string? ExternalMessage { get; set; }

    [JsonPropertyName("externalAudience")]
    public string? ExternalAudience { get; set; }
}

public class MailboxPermissionsDto
{
    [JsonPropertyName("fullAccessPermissions")]
    public List<MailboxPermissionEntryDto> FullAccessPermissions { get; set; } = new();

    [JsonPropertyName("sendAsPermissions")]
    public List<RecipientPermissionEntryDto> SendAsPermissions { get; set; } = new();

    [JsonPropertyName("sendOnBehalfPermissions")]
    public List<SendOnBehalfPermissionEntryDto> SendOnBehalfPermissions { get; set; } = new();
}

public class MailboxPermissionEntryDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("user")]
    public string User { get; set; } = string.Empty;

    [JsonPropertyName("accessRights")]
    public List<string> AccessRights { get; set; } = new();

    [JsonPropertyName("isInherited")]
    public bool IsInherited { get; set; }

    [JsonPropertyName("deny")]
    public bool Deny { get; set; }

    [JsonPropertyName("inheritanceType")]
    public string? InheritanceType { get; set; }

    [JsonPropertyName("autoMapping")]
    public bool? AutoMapping { get; set; }
}

public class RecipientPermissionEntryDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("trustee")]
    public string Trustee { get; set; } = string.Empty;

    [JsonPropertyName("accessControlType")]
    public string AccessControlType { get; set; } = string.Empty;

    [JsonPropertyName("accessRights")]
    public List<string> AccessRights { get; set; } = new();

    [JsonPropertyName("isInherited")]
    public bool IsInherited { get; set; }
}

public class SendOnBehalfPermissionEntryDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;
}

public class GetMailboxesRequest
{
    [JsonPropertyName("recipientTypeDetails")]
    public string? RecipientTypeDetails { get; set; }

    [JsonPropertyName("searchQuery")]
    public string? SearchQuery { get; set; }

    [JsonPropertyName("filter")]
    public string? Filter { get; set; }

    [JsonPropertyName("pageSize")]
    public int PageSize { get; set; } = 50;

    [JsonPropertyName("skip")]
    public int Skip { get; set; }

    [JsonPropertyName("sortBy")]
    public string? SortBy { get; set; }

    [JsonPropertyName("sortDescending")]
    public bool SortDescending { get; set; }
}

public class GetMailboxesResponse
{
    [JsonPropertyName("mailboxes")]
    public List<MailboxListItemDto> Mailboxes { get; set; } = new();

    [JsonPropertyName("totalCount")]
    public int TotalCount { get; set; }

    [JsonPropertyName("skip")]
    public int Skip { get; set; }

    [JsonPropertyName("pageSize")]
    public int PageSize { get; set; }

    [JsonPropertyName("hasMore")]
    public bool HasMore { get; set; }

    [JsonPropertyName("searchQuery")]
    public string? SearchQuery { get; set; }
}

public class MailboxSpaceItemDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;

    [JsonPropertyName("primarySmtpAddress")]
    public string PrimarySmtpAddress { get; set; } = string.Empty;

    [JsonPropertyName("totalItemSize")]
    public string? TotalItemSize { get; set; }

    [JsonPropertyName("totalItemSizeBytes")]
    public long? TotalItemSizeBytes { get; set; }

    [JsonPropertyName("prohibitSendQuota")]
    public string? ProhibitSendQuota { get; set; }

    [JsonPropertyName("prohibitSendQuotaBytes")]
    public long? ProhibitSendQuotaBytes { get; set; }

    [JsonPropertyName("prohibitSendReceiveQuota")]
    public string? ProhibitSendReceiveQuota { get; set; }

    [JsonPropertyName("prohibitSendReceiveQuotaBytes")]
    public long? ProhibitSendReceiveQuotaBytes { get; set; }

    [JsonPropertyName("issueWarningQuota")]
    public string? IssueWarningQuota { get; set; }

    [JsonPropertyName("issueWarningQuotaBytes")]
    public long? IssueWarningQuotaBytes { get; set; }
}

public class GetMailboxSpaceReportRequest
{
}

public class GetMailboxSpaceReportResponse
{
    [JsonPropertyName("mailboxes")]
    public List<MailboxSpaceItemDto> Mailboxes { get; set; } = new();
}

public class GetMailboxDetailsRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("includeStatistics")]
    public bool IncludeStatistics { get; set; } = true;

    [JsonPropertyName("includeRules")]
    public bool IncludeRules { get; set; } = true;

    [JsonPropertyName("includeAutoReply")]
    public bool IncludeAutoReply { get; set; } = true;

    [JsonPropertyName("includePermissions")]
    public bool IncludePermissions { get; set; } = true;
}

public class GetRetentionPoliciesRequest
{
}

public class GetRetentionPoliciesResponse
{
    [JsonPropertyName("policies")]
    public List<RetentionPolicySummaryDto> Policies { get; set; } = new();
}

public class SetRetentionPolicyRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("policyName")]
    public string? PolicyName { get; set; }
}

public class GetMailboxPermissionsRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;
}

public class SetMailboxPermissionRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("user")]
    public string User { get; set; } = string.Empty;

    [JsonPropertyName("permissionType")]
    public PermissionType PermissionType { get; set; }

    [JsonPropertyName("action")]
    public PermissionAction Action { get; set; }

    [JsonPropertyName("autoMapping")]
    public bool? AutoMapping { get; set; }
}

public enum PermissionType
{
    FullAccess,
    SendAs,
    SendOnBehalf
}

public enum PermissionAction
{
    Add,
    Remove,
    Modify
}

public class PermissionsDeltaPlanDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("actions")]
    public List<PermissionDeltaActionDto> Actions { get; set; } = new();

    [JsonPropertyName("totalAdd")]
    public int TotalAdd { get; set; }

    [JsonPropertyName("totalRemove")]
    public int TotalRemove { get; set; }

    [JsonPropertyName("hasChanges")]
    public bool HasChanges => Actions.Count > 0;
}

public class PermissionDeltaActionDto
{
    [JsonPropertyName("action")]
    public PermissionAction Action { get; set; }

    [JsonPropertyName("permissionType")]
    public PermissionType PermissionType { get; set; }

    [JsonPropertyName("user")]
    public string User { get; set; } = string.Empty;

    [JsonPropertyName("autoMapping")]
    public bool? AutoMapping { get; set; }

    [JsonPropertyName("description")]
    public string Description { get; set; } = string.Empty;
}

public class ApplyPermissionsDeltaPlanRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("actions")]
    public List<PermissionDeltaActionDto> Actions { get; set; } = new();
}

public class ApplyPermissionsDeltaPlanResponse
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("totalActions")]
    public int TotalActions { get; set; }

    [JsonPropertyName("successfulActions")]
    public int SuccessfulActions { get; set; }

    [JsonPropertyName("failedActions")]
    public int FailedActions { get; set; }

    [JsonPropertyName("results")]
    public List<PermissionActionResultDto> Results { get; set; } = new();
}

public class PermissionActionResultDto
{
    [JsonPropertyName("action")]
    public PermissionDeltaActionDto Action { get; set; } = new();

    [JsonPropertyName("success")]
    public bool Success { get; set; }

    [JsonPropertyName("errorMessage")]
    public string? ErrorMessage { get; set; }
}

public class SetMailboxFeatureRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("feature")]
    public MailboxFeature Feature { get; set; }

    [JsonPropertyName("enabled")]
    public bool Enabled { get; set; }

    [JsonPropertyName("parameters")]
    public Dictionary<string, object>? Parameters { get; set; }
}

public class UpdateMailboxSettingsRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("archiveEnabled")]
    public bool? ArchiveEnabled { get; set; }

    [JsonPropertyName("litigationHoldEnabled")]
    public bool? LitigationHoldEnabled { get; set; }

    [JsonPropertyName("auditEnabled")]
    public bool? AuditEnabled { get; set; }

    [JsonPropertyName("singleItemRecoveryEnabled")]
    public bool? SingleItemRecoveryEnabled { get; set; }

    [JsonPropertyName("retentionHoldEnabled")]
    public bool? RetentionHoldEnabled { get; set; }

    [JsonPropertyName("forwardingAddress")]
    public string? ForwardingAddress { get; set; }

    [JsonPropertyName("forwardingSmtpAddress")]
    public string? ForwardingSmtpAddress { get; set; }

    [JsonPropertyName("deliverToMailboxAndForward")]
    public bool? DeliverToMailboxAndForward { get; set; }

    [JsonPropertyName("maxSendSize")]
    public string? MaxSendSize { get; set; }

    [JsonPropertyName("maxReceiveSize")]
    public string? MaxReceiveSize { get; set; }
}

public class SetMailboxAutoReplyConfigurationRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("autoReplyState")]
    public string AutoReplyState { get; set; } = "Disabled";

    [JsonPropertyName("startTime")]
    public DateTime? StartTime { get; set; }

    [JsonPropertyName("endTime")]
    public DateTime? EndTime { get; set; }

    [JsonPropertyName("internalMessage")]
    public string? InternalMessage { get; set; }

    [JsonPropertyName("externalMessage")]
    public string? ExternalMessage { get; set; }

    [JsonPropertyName("externalAudience")]
    public string? ExternalAudience { get; set; }
}

public class ConvertMailboxToSharedRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;
}

public class ConvertMailboxToRegularRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;
}

public class RestoreMailboxRequest
{
    [JsonPropertyName("sourceIdentity")]
    public string SourceIdentity { get; set; } = string.Empty;

    [JsonPropertyName("targetMailbox")]
    public string? TargetMailbox { get; set; }

    [JsonPropertyName("allowLegacyDnMismatch")]
    public bool AllowLegacyDnMismatch { get; set; }
}

public class RestoreMailboxResponse
{
    [JsonPropertyName("sourceIdentity")]
    public string SourceIdentity { get; set; } = string.Empty;

    [JsonPropertyName("targetMailbox")]
    public string? TargetMailbox { get; set; }

    [JsonPropertyName("scenario")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public RestoreMailboxScenario Scenario { get; set; }

    [JsonPropertyName("action")]
    public string? Action { get; set; }

    [JsonPropertyName("status")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public RestoreMailboxStatus Status { get; set; }

    [JsonPropertyName("statusDetail")]
    public string? StatusDetail { get; set; }

    [JsonPropertyName("percentComplete")]
    public int? PercentComplete { get; set; }

    [JsonPropertyName("requestGuid")]
    public string? RequestGuid { get; set; }

    [JsonPropertyName("error")]
    public RestoreMailboxErrorDto? Error { get; set; }
}

public class RestoreMailboxErrorDto
{
    [JsonPropertyName("code")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public RestoreMailboxErrorCode Code { get; set; }

    [JsonPropertyName("message")]
    public string Message { get; set; } = string.Empty;
}

public enum RestoreMailboxScenario
{
    Unknown,
    SoftDeleted,
    HardDeleted,
    Existing,
    NotFound
}

public enum RestoreMailboxStatus
{
    NotStarted,
    InProgress,
    Completed,
    Failed
}

public enum RestoreMailboxErrorCode
{
    None,
    UserNotFound,
    LegacyDnMismatch,
    PermissionDenied,
    TargetMailboxRequired,
    Unknown
}

public enum MailboxFeature
{
    Archive,
    LitigationHold,
    Audit,
    SingleItemRecovery,
    RetentionHold
}
