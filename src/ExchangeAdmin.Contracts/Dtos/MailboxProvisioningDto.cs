using System.Text.Json.Serialization;
using ExchangeAdmin.Contracts.Paging;

namespace ExchangeAdmin.Contracts.Dtos;

public sealed class MailboxProvisioningCandidateDto
{
    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;

    [JsonPropertyName("userPrincipalName")]
    public string UserPrincipalName { get; set; } = string.Empty;

    [JsonPropertyName("mail")]
    public string? Mail { get; set; }

    [JsonPropertyName("accountEnabled")]
    public bool AccountEnabled { get; set; }

    [JsonPropertyName("hasAssignedLicense")]
    public bool HasAssignedLicense { get; set; }

    [JsonPropertyName("hasMailAddress")]
    public bool HasMailAddress { get; set; }

    [JsonPropertyName("assignedLicenses")]
    public List<UserLicenseDto> AssignedLicenses { get; set; } = new();
}

public sealed class GetMailboxProvisioningCandidatesRequest
{
    [JsonPropertyName("searchQuery")]
    public string? SearchQuery { get; set; }

    [JsonPropertyName("onlyWithoutLicense")]
    public bool OnlyWithoutLicense { get; set; } = true;

    [JsonPropertyName("onlyWithoutMail")]
    public bool OnlyWithoutMail { get; set; } = true;

    [JsonPropertyName("pageSize")]
    public int PageSize { get; set; } = PagingDefaults.DefaultPageSize;

    [JsonPropertyName("skip")]
    public int Skip { get; set; }
}

public sealed class GetMailboxProvisioningCandidatesResponse
{
    [JsonPropertyName("candidates")]
    public List<MailboxProvisioningCandidateDto> Candidates { get; set; } = new();

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
