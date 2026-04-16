using System.Text.Json.Serialization;
using ExchangeAdmin.Contracts.Paging;

namespace ExchangeAdmin.Contracts.Dtos;

public class MobileDeviceListItemDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("guid")]
    public string? Guid { get; set; }

    [JsonPropertyName("deviceId")]
    public string DeviceId { get; set; } = string.Empty;

    [JsonPropertyName("friendlyName")]
    public string? FriendlyName { get; set; }

    [JsonPropertyName("deviceType")]
    public string? DeviceType { get; set; }

    [JsonPropertyName("deviceModel")]
    public string? DeviceModel { get; set; }

    [JsonPropertyName("deviceUserAgent")]
    public string? DeviceUserAgent { get; set; }

    [JsonPropertyName("deviceOS")]
    public string? DeviceOS { get; set; }

    [JsonPropertyName("clientType")]
    public string? ClientType { get; set; }

    [JsonPropertyName("userDisplayName")]
    public string? UserDisplayName { get; set; }

    [JsonPropertyName("userPrincipalName")]
    public string? UserPrincipalName { get; set; }

    [JsonPropertyName("mailboxIdentity")]
    public string MailboxIdentity { get; set; } = string.Empty;

    [JsonPropertyName("mailboxDisplayName")]
    public string MailboxDisplayName { get; set; } = string.Empty;

    [JsonPropertyName("currentMailboxPolicy")]
    public string? CurrentMailboxPolicy { get; set; }

    [JsonPropertyName("deviceAccessState")]
    public string DeviceAccessState { get; set; } = string.Empty;

    [JsonPropertyName("deviceAccessStateReason")]
    public string? DeviceAccessStateReason { get; set; }

    [JsonPropertyName("firstSyncTime")]
    public DateTime? FirstSyncTime { get; set; }

    [JsonPropertyName("lastSuccessSync")]
    public DateTime? LastSuccessSync { get; set; }

    [JsonPropertyName("lastPolicyUpdateTime")]
    public DateTime? LastPolicyUpdateTime { get; set; }

    [JsonPropertyName("status")]
    public string? Status { get; set; }
}

public class MobileDeviceMailboxPolicyDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("isDefault")]
    public bool IsDefault { get; set; }

    [JsonPropertyName("passwordEnabled")]
    public bool? PasswordEnabled { get; set; }

    [JsonPropertyName("alphanumericPasswordRequired")]
    public bool? AlphanumericPasswordRequired { get; set; }

    [JsonPropertyName("allowNonProvisionableDevices")]
    public bool? AllowNonProvisionableDevices { get; set; }

    [JsonPropertyName("deviceEncryptionEnabled")]
    public bool? DeviceEncryptionEnabled { get; set; }

    [JsonPropertyName("attachmentsEnabled")]
    public bool? AttachmentsEnabled { get; set; }

    [JsonPropertyName("maxAttachmentSize")]
    public string? MaxAttachmentSize { get; set; }
}

public class GetMobileDevicesRequest
{
    [JsonPropertyName("searchQuery")]
    public string? SearchQuery { get; set; }

    [JsonPropertyName("accessState")]
    public string? AccessState { get; set; }

    [JsonPropertyName("pageSize")]
    public int PageSize { get; set; } = PagingDefaults.DefaultPageSize;

    [JsonPropertyName("skip")]
    public int Skip { get; set; }

    [JsonPropertyName("sortBy")]
    public string? SortBy { get; set; }

    [JsonPropertyName("sortDescending")]
    public bool SortDescending { get; set; }
}

public class GetMobileDeviceDetailsRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("mailboxIdentity")]
    public string? MailboxIdentity { get; set; }
}

public class GetMobileDevicesResponse
{
    [JsonPropertyName("devices")]
    public List<MobileDeviceListItemDto> Devices { get; set; } = new();

    [JsonPropertyName("totalCount")]
    public int TotalCount { get; set; }

    [JsonPropertyName("skip")]
    public int Skip { get; set; }

    [JsonPropertyName("pageSize")]
    public int PageSize { get; set; }

    [JsonPropertyName("hasMore")]
    public bool HasMore { get; set; }

    [JsonPropertyName("isTotalCountExact")]
    public bool IsTotalCountExact { get; set; } = true;

    [JsonPropertyName("searchQuery")]
    public string? SearchQuery { get; set; }
}

public class GetMobileDeviceDetailsResponse
{
    [JsonPropertyName("device")]
    public MobileDeviceListItemDto? Device { get; set; }
}

public class GetMobileDeviceMailboxPoliciesRequest
{
}

public class GetMobileDeviceMailboxPoliciesResponse
{
    [JsonPropertyName("policies")]
    public List<MobileDeviceMailboxPolicyDto> Policies { get; set; } = new();
}

public class SetMobileDeviceAccessStateRequest
{
    [JsonPropertyName("mailboxIdentity")]
    public string MailboxIdentity { get; set; } = string.Empty;

    [JsonPropertyName("deviceId")]
    public string DeviceId { get; set; } = string.Empty;

    [JsonPropertyName("accessState")]
    public string AccessState { get; set; } = string.Empty;
}

public class ClearMobileDeviceRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;
}

public class SetMobileDeviceMailboxPolicyRequest
{
    [JsonPropertyName("mailboxIdentity")]
    public string MailboxIdentity { get; set; } = string.Empty;

    [JsonPropertyName("policyIdentity")]
    public string? PolicyIdentity { get; set; }
}
