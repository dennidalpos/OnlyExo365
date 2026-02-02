using System.Text.Json.Serialization;

namespace ExchangeAdmin.Contracts.Dtos;

public class DashboardStatsDto
{
    [JsonPropertyName("retrievedAt")]
    public DateTime RetrievedAt { get; set; } = DateTime.UtcNow;

    [JsonPropertyName("mailboxCounts")]
    public MailboxCountsDto MailboxCounts { get; set; } = new();

    [JsonPropertyName("groupCounts")]
    public GroupCountsDto GroupCounts { get; set; } = new();

    [JsonPropertyName("warnings")]
    public List<string> Warnings { get; set; } = new();

    [JsonPropertyName("isLargeTenant")]
    public bool IsLargeTenant { get; set; }

    [JsonPropertyName("largeTenantThreshold")]
    public int LargeTenantThreshold { get; set; } = 1000;

    [JsonPropertyName("licenses")]
    public List<TenantLicenseDto> Licenses { get; set; } = new();

    [JsonPropertyName("adminUsers")]
    public List<AdminRoleMemberDto> AdminUsers { get; set; } = new();
}

public class MailboxCountsDto
{
    [JsonPropertyName("userMailboxes")]
    public int UserMailboxes { get; set; }

    [JsonPropertyName("sharedMailboxes")]
    public int SharedMailboxes { get; set; }

    [JsonPropertyName("roomMailboxes")]
    public int RoomMailboxes { get; set; }

    [JsonPropertyName("equipmentMailboxes")]
    public int EquipmentMailboxes { get; set; }

    [JsonPropertyName("total")]
    public int Total => UserMailboxes + SharedMailboxes + RoomMailboxes + EquipmentMailboxes;

    [JsonPropertyName("isApproximate")]
    public bool IsApproximate { get; set; }
}

public class GroupCountsDto
{
    [JsonPropertyName("distributionGroups")]
    public int DistributionGroups { get; set; }

    [JsonPropertyName("dynamicDistributionGroups")]
    public int DynamicDistributionGroups { get; set; }

    [JsonPropertyName("unifiedGroups")]
    public int? UnifiedGroups { get; set; }

    [JsonPropertyName("unifiedGroupsAvailable")]
    public bool UnifiedGroupsAvailable { get; set; }

    [JsonPropertyName("total")]
    public int Total => DistributionGroups + DynamicDistributionGroups + (UnifiedGroups ?? 0);
}

public class GetDashboardStatsRequest
{
    [JsonPropertyName("includeUnifiedGroups")]
    public bool IncludeUnifiedGroups { get; set; } = true;

    [JsonPropertyName("quickCount")]
    public bool QuickCount { get; set; } = true;
}
