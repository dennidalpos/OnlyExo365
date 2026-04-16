using System.Text.Json.Serialization;
using ExchangeAdmin.Contracts.Paging;

namespace ExchangeAdmin.Contracts.Dtos;

public class ResourceMailboxListItemDto
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;

    [JsonPropertyName("alias")]
    public string Alias { get; set; } = string.Empty;

    [JsonPropertyName("primarySmtpAddress")]
    public string PrimarySmtpAddress { get; set; } = string.Empty;

    [JsonPropertyName("resourceType")]
    public string ResourceType { get; set; } = "Room";

    [JsonPropertyName("recipientTypeDetails")]
    public string RecipientTypeDetails { get; set; } = string.Empty;

    [JsonPropertyName("hiddenFromAddressListsEnabled")]
    public bool HiddenFromAddressListsEnabled { get; set; }
}

public class ResourceMailboxDetailsDto : ResourceMailboxListItemDto
{
    [JsonPropertyName("name")]
    public string? Name { get; set; }

    [JsonPropertyName("whenCreated")]
    public DateTime? WhenCreated { get; set; }

    [JsonPropertyName("bookingSettings")]
    public ResourceBookingSettingsDto BookingSettings { get; set; } = new();

    [JsonPropertyName("permissions")]
    public MailboxPermissionsDto Permissions { get; set; } = new();
}

public class ResourceBookingSettingsDto
{
    [JsonPropertyName("automateProcessing")]
    public string AutomateProcessing { get; set; } = "AutoAccept";

    [JsonPropertyName("allowConflicts")]
    public bool AllowConflicts { get; set; }

    [JsonPropertyName("allBookInPolicy")]
    public bool AllBookInPolicy { get; set; }

    [JsonPropertyName("allRequestInPolicy")]
    public bool AllRequestInPolicy { get; set; }

    [JsonPropertyName("allRequestOutOfPolicy")]
    public bool AllRequestOutOfPolicy { get; set; }

    [JsonPropertyName("bookingWindowInDays")]
    public int? BookingWindowInDays { get; set; }

    [JsonPropertyName("maximumDurationInMinutes")]
    public int? MaximumDurationInMinutes { get; set; }

    [JsonPropertyName("deleteSubject")]
    public bool? DeleteSubject { get; set; }

    [JsonPropertyName("addOrganizerToSubject")]
    public bool? AddOrganizerToSubject { get; set; }

    [JsonPropertyName("removePrivateProperty")]
    public bool? RemovePrivateProperty { get; set; }

    [JsonPropertyName("enforceSchedulingHorizon")]
    public bool? EnforceSchedulingHorizon { get; set; }

    [JsonPropertyName("bookInPolicy")]
    public List<string> BookInPolicy { get; set; } = new();

    [JsonPropertyName("requestInPolicy")]
    public List<string> RequestInPolicy { get; set; } = new();

    [JsonPropertyName("requestOutOfPolicy")]
    public List<string> RequestOutOfPolicy { get; set; } = new();

    [JsonPropertyName("resourceDelegates")]
    public List<string> ResourceDelegates { get; set; } = new();
}

public class GetResourceMailboxesRequest
{
    [JsonPropertyName("resourceType")]
    public string? ResourceType { get; set; }

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

public class GetResourceMailboxesResponse
{
    [JsonPropertyName("resources")]
    public List<ResourceMailboxListItemDto> Resources { get; set; } = new();

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

public class GetResourceMailboxDetailsRequest
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;
}

public class UpsertResourceMailboxRequest
{
    [JsonPropertyName("identity")]
    public string? Identity { get; set; }

    [JsonPropertyName("resourceType")]
    public string ResourceType { get; set; } = "Room";

    [JsonPropertyName("displayName")]
    public string DisplayName { get; set; } = string.Empty;

    [JsonPropertyName("name")]
    public string? Name { get; set; }

    [JsonPropertyName("alias")]
    public string Alias { get; set; } = string.Empty;

    [JsonPropertyName("primarySmtpAddress")]
    public string PrimarySmtpAddress { get; set; } = string.Empty;

    [JsonPropertyName("hiddenFromAddressListsEnabled")]
    public bool HiddenFromAddressListsEnabled { get; set; }

    [JsonPropertyName("bookingSettings")]
    public ResourceBookingSettingsDto BookingSettings { get; set; } = new();
}

public class UpsertResourceMailboxResponse
{
    [JsonPropertyName("identity")]
    public string Identity { get; set; } = string.Empty;

    [JsonPropertyName("primarySmtpAddress")]
    public string PrimarySmtpAddress { get; set; } = string.Empty;
}
