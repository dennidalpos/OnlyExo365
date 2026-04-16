using System.Text.Json.Serialization;
using ExchangeAdmin.Contracts.Paging;

namespace ExchangeAdmin.Contracts.Dtos;

public class PublicFolderListItemDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("parentPath")]
    public string ParentPath { get; set; } = "\\";

    [JsonPropertyName("mailEnabled")]
    public bool MailEnabled { get; set; }

    [JsonPropertyName("alias")]
    public string? Alias { get; set; }

    [JsonPropertyName("primarySmtpAddress")]
    public string? PrimarySmtpAddress { get; set; }

    [JsonPropertyName("hiddenFromAddressListsEnabled")]
    public bool HiddenFromAddressListsEnabled { get; set; }

    [JsonPropertyName("hasSubFolders")]
    public bool HasSubFolders { get; set; }
}

public class PublicFolderPermissionEntryDto
{
    [JsonPropertyName("user")]
    public string User { get; set; } = string.Empty;

    [JsonPropertyName("accessRights")]
    public List<string> AccessRights { get; set; } = new();

    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;
}

public class PublicFolderDetailsDto : PublicFolderListItemDto
{
    [JsonPropertyName("itemCount")]
    public int? ItemCount { get; set; }

    [JsonPropertyName("totalItemSize")]
    public string? TotalItemSize { get; set; }

    [JsonPropertyName("permissions")]
    public List<PublicFolderPermissionEntryDto> Permissions { get; set; } = new();
}

public class GetPublicFoldersRequest
{
    [JsonPropertyName("searchQuery")]
    public string? SearchQuery { get; set; }

    [JsonPropertyName("mailEnabledOnly")]
    public bool? MailEnabledOnly { get; set; }

    [JsonPropertyName("pageSize")]
    public int PageSize { get; set; } = PagingDefaults.DefaultPageSize;

    [JsonPropertyName("skip")]
    public int Skip { get; set; }

    [JsonPropertyName("sortBy")]
    public string? SortBy { get; set; }

    [JsonPropertyName("sortDescending")]
    public bool SortDescending { get; set; }
}

public class GetPublicFoldersResponse
{
    [JsonPropertyName("folders")]
    public List<PublicFolderListItemDto> Folders { get; set; } = new();

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

public class GetPublicFolderDetailsRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;
}

public class UpsertPublicFolderRequest
{
    [JsonPropertyName("identity")]
    public string? Identity { get; set; }

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("parentPath")]
    public string? ParentPath { get; set; }

    [JsonPropertyName("mailEnabled")]
    public bool MailEnabled { get; set; }

    [JsonPropertyName("alias")]
    public string? Alias { get; set; }

    [JsonPropertyName("primarySmtpAddress")]
    public string? PrimarySmtpAddress { get; set; }

    [JsonPropertyName("hiddenFromAddressListsEnabled")]
    public bool HiddenFromAddressListsEnabled { get; set; }
}

public class UpsertPublicFolderResponse
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("mailEnabled")]
    public bool MailEnabled { get; set; }

    [JsonPropertyName("primarySmtpAddress")]
    public string? PrimarySmtpAddress { get; set; }
}

public class SetPublicFolderClientPermissionRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("user")]
    public string User { get; set; } = string.Empty;

    [JsonPropertyName("action")]
    public PermissionAction Action { get; set; }

    [JsonPropertyName("accessRights")]
    public List<string> AccessRights { get; set; } = new();
}

public class RemovePublicFolderRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("recursive")]
    public bool Recursive { get; set; }
}
