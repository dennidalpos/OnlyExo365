using System.Text.Json.Serialization;
using ExchangeAdmin.Contracts.Paging;

namespace ExchangeAdmin.Contracts.Dtos;

public class GetDeletedMailboxesRequest
{
    [JsonPropertyName("searchQuery")]
    public string? SearchQuery { get; set; }

    [JsonPropertyName("includeSoftDeleted")]
    public bool IncludeSoftDeleted { get; set; } = true;

    [JsonPropertyName("includeInactive")]
    public bool IncludeInactive { get; set; } = true;

    [JsonPropertyName("pageSize")]
    public int PageSize { get; set; } = PagingDefaults.DefaultPageSize;

    [JsonPropertyName("skip")]
    public int Skip { get; set; }
}

public class GetDeletedMailboxesResponse
{
    [JsonPropertyName("mailboxes")]
    public List<DeletedMailboxItemDto> Mailboxes { get; set; } = new();

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

public class DeletedMailboxItemDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;

    [JsonPropertyName("userPrincipalName")]
    public string? UserPrincipalName { get; set; }

    [JsonPropertyName("primarySmtpAddress")]
    public string PrimarySmtpAddress { get; set; } = string.Empty;

    [JsonPropertyName("recipientTypeDetails")]
    public string RecipientTypeDetails { get; set; } = string.Empty;

    [JsonPropertyName("alias")]
    public string? Alias { get; set; }

    [JsonPropertyName("deletionType")]
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public DeletedMailboxDeletionType DeletionType { get; set; }
}

public enum DeletedMailboxDeletionType
{
    Unknown,
    SoftDeleted,
    Inactive,
    HardDeleted
}
