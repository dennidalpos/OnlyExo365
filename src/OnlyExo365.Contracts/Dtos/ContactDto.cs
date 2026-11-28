using System.Text.Json.Serialization;
using OnlyExo365.Contracts.Paging;
using OnlyExo365.Contracts.Security;

namespace OnlyExo365.Contracts.Dtos;

public class ContactListItemDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("guid")]
    public string? Guid { get; set; }

    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;

    [JsonPropertyName("name")]
    public string? Name { get; set; }

    [JsonPropertyName("alias")]
    public string? Alias { get; set; }

    [JsonPropertyName("primarySmtpAddress")]
    public string PrimarySmtpAddress { get; set; } = string.Empty;

    [JsonPropertyName("externalEmailAddress")]
    public string? ExternalEmailAddress { get; set; }

    [JsonPropertyName("recipientType")]
    public string RecipientType { get; set; } = string.Empty;

    [JsonPropertyName("recipientTypeDetails")]
    public string RecipientTypeDetails { get; set; } = string.Empty;

    [JsonPropertyName("contactKind")]
    public string ContactKind { get; set; } = "MailContact";

    [JsonPropertyName("userPrincipalName")]
    public string? UserPrincipalName { get; set; }

    [JsonPropertyName("hiddenFromAddressListsEnabled")]
    public bool HiddenFromAddressListsEnabled { get; set; }
}

public class ContactDetailsDto : ContactListItemDto
{
    [JsonPropertyName("whenCreated")]
    public DateTime? WhenCreated { get; set; }

    [JsonPropertyName("emailAddresses")]
    public List<string> EmailAddresses { get; set; } = new();
}

public class GetContactsRequest
{
    [JsonPropertyName("contactKind")]
    public string? ContactKind { get; set; }

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

public class GetContactsResponse
{
    [JsonPropertyName("contacts")]
    public List<ContactListItemDto> Contacts { get; set; } = new();

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

public class GetContactDetailsRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("contactKind")]
    public string? ContactKind { get; set; }
}

public class UpsertContactRequest
{
    [JsonPropertyName("identity")]
    public string? Identity { get; set; }

    [JsonPropertyName("contactKind")]
    public string ContactKind { get; set; } = "MailContact";

    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;

    [JsonPropertyName("name")]
    public string? Name { get; set; }

    [JsonPropertyName("alias")]
    public string Alias { get; set; } = string.Empty;

    [JsonPropertyName("primarySmtpAddress")]
    public string PrimarySmtpAddress { get; set; } = string.Empty;

    [JsonPropertyName("externalEmailAddress")]
    public string ExternalEmailAddress { get; set; } = string.Empty;

    [JsonPropertyName("userPrincipalName")]
    public string? UserPrincipalName { get; set; }

    [JsonIgnore]
    public string? Password { get; set; }

    [JsonPropertyName("passwordSecret")]
    public ProtectedSecretReference? PasswordSecret { get; set; }

    [JsonPropertyName("hiddenFromAddressListsEnabled")]
    public bool HiddenFromAddressListsEnabled { get; set; }
}

public class RemoveContactRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("contactKind")]
    public string ContactKind { get; set; } = "MailContact";
}

