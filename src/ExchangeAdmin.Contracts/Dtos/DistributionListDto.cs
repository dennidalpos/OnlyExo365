using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Dtos;

/// <summary>
/// Distribution list item (lightweight).
/// </summary>
public class DistributionListItemDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("guid")]
    public string? Guid { get; set; }

    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;

    [JsonPropertyName("primarySmtpAddress")]
    public string PrimarySmtpAddress { get; set; } = string.Empty;

    [JsonPropertyName("alias")]
    public string? Alias { get; set; }

    [JsonPropertyName("groupType")]
    public string GroupType { get; set; } = string.Empty;

    [JsonPropertyName("recipientType")]
    public string RecipientType { get; set; } = string.Empty;

    [JsonPropertyName("recipientTypeDetails")]
    public string RecipientTypeDetails { get; set; } = string.Empty;

    [JsonPropertyName("isDynamic")]
    public bool IsDynamic { get; set; }

    [JsonPropertyName("managedBy")]
    public List<string> ManagedBy { get; set; } = new();

    [JsonPropertyName("memberCount")]
    public int? MemberCount { get; set; }
}

/// <summary>
/// Full distribution list DTO (backward compatibility).
/// </summary>
public class DistributionListDto : DistributionListItemDto
{
    [JsonPropertyName("whenCreated")]
    public DateTime? WhenCreated { get; set; }

    [JsonPropertyName("whenChanged")]
    public DateTime? WhenChanged { get; set; }
}

/// <summary>
/// Distribution list details (full information).
/// </summary>
public class DistributionListDetailsDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("guid")]
    public string? Guid { get; set; }

    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;

    [JsonPropertyName("primarySmtpAddress")]
    public string PrimarySmtpAddress { get; set; } = string.Empty;

    [JsonPropertyName("alias")]
    public string? Alias { get; set; }

    [JsonPropertyName("groupType")]
    public string GroupType { get; set; } = string.Empty;

    [JsonPropertyName("recipientType")]
    public string RecipientType { get; set; } = string.Empty;

    [JsonPropertyName("recipientTypeDetails")]
    public string RecipientTypeDetails { get; set; } = string.Empty;

    [JsonPropertyName("emailAddresses")]
    public List<string> EmailAddresses { get; set; } = new();

    [JsonPropertyName("managedBy")]
    public List<string> ManagedBy { get; set; } = new();

    [JsonPropertyName("acceptMessagesOnlyFrom")]
    public List<string> AcceptMessagesOnlyFrom { get; set; } = new();

    [JsonPropertyName("rejectMessagesFrom")]
    public List<string> RejectMessagesFrom { get; set; } = new();

    [JsonPropertyName("requireSenderAuthenticationEnabled")]
    public bool RequireSenderAuthenticationEnabled { get; set; }

    [JsonPropertyName("hiddenFromAddressListsEnabled")]
    public bool HiddenFromAddressListsEnabled { get; set; }

    [JsonPropertyName("memberJoinRestriction")]
    public string? MemberJoinRestriction { get; set; }

    [JsonPropertyName("memberDepartRestriction")]
    public string? MemberDepartRestriction { get; set; }

    [JsonPropertyName("whenCreated")]
    public DateTime? WhenCreated { get; set; }

    [JsonPropertyName("whenChanged")]
    public DateTime? WhenChanged { get; set; }

    // Members (loaded with paging)
    [JsonPropertyName("members")]
    public GroupMembersPageDto? Members { get; set; }
}

/// <summary>
/// Dynamic distribution list details.
/// </summary>
public class DynamicDistributionListDetailsDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("guid")]
    public string? Guid { get; set; }

    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;

    [JsonPropertyName("primarySmtpAddress")]
    public string PrimarySmtpAddress { get; set; } = string.Empty;

    [JsonPropertyName("alias")]
    public string? Alias { get; set; }

    [JsonPropertyName("recipientFilter")]
    public string? RecipientFilter { get; set; }

    [JsonPropertyName("includedRecipients")]
    public string? IncludedRecipients { get; set; }

    [JsonPropertyName("conditionalDepartment")]
    public List<string>? ConditionalDepartment { get; set; }

    [JsonPropertyName("conditionalCompany")]
    public List<string>? ConditionalCompany { get; set; }

    [JsonPropertyName("conditionalStateOrProvince")]
    public List<string>? ConditionalStateOrProvince { get; set; }

    [JsonPropertyName("conditionalCustomAttribute1")]
    public List<string>? ConditionalCustomAttribute1 { get; set; }

    [JsonPropertyName("managedBy")]
    public List<string> ManagedBy { get; set; } = new();

    [JsonPropertyName("emailAddresses")]
    public List<string> EmailAddresses { get; set; } = new();

    [JsonPropertyName("whenCreated")]
    public DateTime? WhenCreated { get; set; }

    [JsonPropertyName("whenChanged")]
    public DateTime? WhenChanged { get; set; }

    // Preview members (loaded on demand)
    [JsonPropertyName("previewMembers")]
    public GroupMembersPageDto? PreviewMembers { get; set; }

    [JsonPropertyName("previewWarning")]
    public string? PreviewWarning { get; set; }
}

/// <summary>
/// Unified group (Microsoft 365 Group) details.
/// </summary>
public class UnifiedGroupDetailsDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("guid")]
    public string? Guid { get; set; }

    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;

    [JsonPropertyName("primarySmtpAddress")]
    public string PrimarySmtpAddress { get; set; } = string.Empty;

    [JsonPropertyName("alias")]
    public string? Alias { get; set; }

    [JsonPropertyName("accessType")]
    public string? AccessType { get; set; }

    [JsonPropertyName("classification")]
    public string? Classification { get; set; }

    [JsonPropertyName("managedBy")]
    public List<string> ManagedBy { get; set; } = new();

    [JsonPropertyName("notes")]
    public string? Notes { get; set; }

    [JsonPropertyName("hideFromAddressLists")]
    public bool HideFromAddressLists { get; set; }

    [JsonPropertyName("hideFromExchangeClients")]
    public bool HideFromExchangeClients { get; set; }

    [JsonPropertyName("subscriptionEnabled")]
    public bool SubscriptionEnabled { get; set; }

    [JsonPropertyName("welcomeMessageEnabled")]
    public bool WelcomeMessageEnabled { get; set; }

    [JsonPropertyName("whenCreated")]
    public DateTime? WhenCreated { get; set; }

    [JsonPropertyName("whenChanged")]
    public DateTime? WhenChanged { get; set; }

    // Members (loaded with paging via Get-UnifiedGroupLinks)
    [JsonPropertyName("members")]
    public GroupMembersPageDto? Members { get; set; }

    [JsonPropertyName("owners")]
    public List<GroupMemberDto>? Owners { get; set; }
}

/// <summary>
/// DTO per membro distribution list.
/// </summary>
public class DistributionListMemberDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;

    [JsonPropertyName("primarySmtpAddress")]
    public string PrimarySmtpAddress { get; set; } = string.Empty;

    [JsonPropertyName("recipientType")]
    public string RecipientType { get; set; } = string.Empty;
}

/// <summary>
/// Group member.
/// </summary>
public class GroupMemberDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("primarySmtpAddress")]
    public string? PrimarySmtpAddress { get; set; }

    [JsonPropertyName("recipientType")]
    public string? RecipientType { get; set; }
}

/// <summary>
/// Page of group members.
/// </summary>
public class GroupMembersPageDto
{
    [JsonPropertyName("members")]
    public List<GroupMemberDto> Members { get; set; } = new();

    [JsonPropertyName("totalCount")]
    public int TotalCount { get; set; }

    [JsonPropertyName("skip")]
    public int Skip { get; set; }

    [JsonPropertyName("pageSize")]
    public int PageSize { get; set; }

    [JsonPropertyName("hasMore")]
    public bool HasMore { get; set; }
}

/// <summary>
/// Request distribution list.
/// </summary>
public class GetDistributionListsRequest
{
    [JsonPropertyName("searchQuery")]
    public string? SearchQuery { get; set; }

    [JsonPropertyName("filter")]
    public string? Filter { get; set; }

    [JsonPropertyName("includeDynamic")]
    public bool IncludeDynamic { get; set; } = true;

    [JsonPropertyName("pageSize")]
    public int PageSize { get; set; } = 50;

    [JsonPropertyName("skip")]
    public int Skip { get; set; }

    [JsonPropertyName("sortBy")]
    public string? SortBy { get; set; }

    [JsonPropertyName("sortDescending")]
    public bool SortDescending { get; set; }
}

/// <summary>
/// Response distribution list.
/// </summary>
public class GetDistributionListsResponse
{
    [JsonPropertyName("distributionLists")]
    public List<DistributionListItemDto> DistributionLists { get; set; } = new();

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

/// <summary>
/// Request distribution list details.
/// </summary>
public class GetDistributionListDetailsRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("includeMembers")]
    public bool IncludeMembers { get; set; } = true;

    [JsonPropertyName("membersPageSize")]
    public int MembersPageSize { get; set; } = 50;
}

/// <summary>
/// Request group members.
/// </summary>
public class GetGroupMembersRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("groupType")]
    public string GroupType { get; set; } = "DistributionGroup";

    [JsonPropertyName("pageSize")]
    public int PageSize { get; set; } = 50;

    [JsonPropertyName("skip")]
    public int Skip { get; set; }
}

/// <summary>
/// Request to add/remove group member.
/// </summary>
public class ModifyGroupMemberRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("member")]
    public string Member { get; set; } = string.Empty;

    [JsonPropertyName("action")]
    public GroupMemberAction Action { get; set; }

    [JsonPropertyName("groupType")]
    public string GroupType { get; set; } = "DistributionGroup";
}

/// <summary>
/// Group member action.
/// </summary>
public enum GroupMemberAction
{
    Add,
    Remove
}

/// <summary>
/// Request to preview dynamic distribution group members.
/// </summary>
public class PreviewDynamicGroupMembersRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("maxResults")]
    public int MaxResults { get; set; } = 100;
}

/// <summary>
/// Response for preview dynamic group members.
/// </summary>
public class PreviewDynamicGroupMembersResponse
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("members")]
    public List<GroupMemberDto> Members { get; set; } = new();

    [JsonPropertyName("totalCount")]
    public int TotalCount { get; set; }

    [JsonPropertyName("isLimited")]
    public bool IsLimited { get; set; }

    [JsonPropertyName("warning")]
    public string? Warning { get; set; }
}
