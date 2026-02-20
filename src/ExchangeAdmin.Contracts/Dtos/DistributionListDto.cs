using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Dtos;

             
                                         
              
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

             
                                                        
              
public class DistributionListDto : DistributionListItemDto
{
    [JsonPropertyName("whenCreated")]
    public DateTime? WhenCreated { get; set; }

    [JsonPropertyName("whenChanged")]
    public DateTime? WhenChanged { get; set; }
}

             
                                                 
              
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

                                   
    [JsonPropertyName("members")]
    public GroupMembersPageDto? Members { get; set; }
}

             
                                      
              
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

                                         
    [JsonPropertyName("previewMembers")]
    public GroupMembersPageDto? PreviewMembers { get; set; }

    [JsonPropertyName("previewWarning")]
    public string? PreviewWarning { get; set; }
}

             
                                                
              
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

                                                             
    [JsonPropertyName("members")]
    public GroupMembersPageDto? Members { get; set; }

    [JsonPropertyName("owners")]
    public List<GroupMemberDto>? Owners { get; set; }
}

             
                                     
              
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

             
                                      
              
public class GetDistributionListDetailsRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("includeMembers")]
    public bool IncludeMembers { get; set; } = true;

    [JsonPropertyName("membersPageSize")]
    public int MembersPageSize { get; set; } = 50;
}

             
                          
              
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

             
                                                 
              
public class SetDistributionListSettingsRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("groupType")]
    public string GroupType { get; set; } = "DistributionGroup";

    [JsonPropertyName("requireSenderAuthenticationEnabled")]
    public bool? RequireSenderAuthenticationEnabled { get; set; }

    [JsonPropertyName("acceptMessagesOnlyFrom")]
    public List<string>? AcceptMessagesOnlyFrom { get; set; }

    [JsonPropertyName("rejectMessagesFrom")]
    public List<string>? RejectMessagesFrom { get; set; }
}

public class CreateDistributionListRequest
{
    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;

    [JsonPropertyName("alias")]
    public string Alias { get; set; } = string.Empty;

    [JsonPropertyName("primarySmtpAddress")]
    public string PrimarySmtpAddress { get; set; } = string.Empty;
}

             
                        
              
public enum GroupMemberAction
{
    Add,
    Remove
}

             
                                                          
              
public class PreviewDynamicGroupMembersRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("maxResults")]
    public int MaxResults { get; set; } = 100;
}

             
                                               
              
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
