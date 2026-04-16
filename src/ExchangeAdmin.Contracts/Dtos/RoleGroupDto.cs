using System.Text.Json.Serialization;
using ExchangeAdmin.Contracts.Paging;

namespace ExchangeAdmin.Contracts.Dtos;

public class RoleGroupListItemDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;

    [JsonPropertyName("description")]
    public string? Description { get; set; }

    [JsonPropertyName("memberCount")]
    public int MemberCount { get; set; }

    [JsonPropertyName("roleCount")]
    public int RoleCount { get; set; }

    [JsonPropertyName("managedBy")]
    public List<string> ManagedBy { get; set; } = new();

    [JsonPropertyName("whenCreated")]
    public DateTime? WhenCreated { get; set; }

    [JsonPropertyName("whenChanged")]
    public DateTime? WhenChanged { get; set; }
}

public class RoleGroupDetailsDto : RoleGroupListItemDto
{
    [JsonPropertyName("roles")]
    public List<string> Roles { get; set; } = new();

    [JsonPropertyName("members")]
    public List<RoleGroupMemberDto> Members { get; set; } = new();

    [JsonPropertyName("customRecipientWriteScope")]
    public string? CustomRecipientWriteScope { get; set; }

    [JsonPropertyName("customConfigWriteScope")]
    public string? CustomConfigWriteScope { get; set; }

    [JsonPropertyName("recipientReadScope")]
    public string? RecipientReadScope { get; set; }

    [JsonPropertyName("recipientWriteScope")]
    public string? RecipientWriteScope { get; set; }
}

public class RoleGroupMemberDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;
}

public class GetRoleGroupsRequest
{
    [JsonPropertyName("searchQuery")]
    public string? SearchQuery { get; set; }

    [JsonPropertyName("pageSize")]
    public int PageSize { get; set; } = PagingDefaults.DefaultPageSize;

    [JsonPropertyName("skip")]
    public int Skip { get; set; }

    [JsonPropertyName("sortBy")]
    public string? SortBy { get; set; }

    [JsonPropertyName("sortDescending")]
    public bool SortDescending { get; set; }
}

public class GetRoleGroupsResponse
{
    [JsonPropertyName("roleGroups")]
    public List<RoleGroupListItemDto> RoleGroups { get; set; } = new();

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

public class GetRoleGroupDetailsRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;
}

public class UpsertRoleGroupRequest
{
    [JsonPropertyName("identity")]
    public string? Identity { get; set; }

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("description")]
    public string? Description { get; set; }

    [JsonPropertyName("roles")]
    public List<string> Roles { get; set; } = new();

    [JsonPropertyName("members")]
    public List<string> Members { get; set; } = new();

    [JsonPropertyName("copyFromRoleGroup")]
    public string? CopyFromRoleGroup { get; set; }
}

public class ModifyRoleGroupMemberRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("member")]
    public string Member { get; set; } = string.Empty;

    [JsonPropertyName("action")]
    public RoleGroupMemberAction Action { get; set; }
}

public enum RoleGroupMemberAction
{
    Add,
    Remove
}
