namespace ExchangeAdmin.Contracts.Messages;

/// <summary>
/// Tipi di operazione supportate dal worker.
/// </summary>
public enum OperationType
{
    // Connection management
    ConnectExchangeInteractive,
    DisconnectExchange,
    GetConnectionStatus,

    // Capability detection
    DetectCapabilities,

    // Demo/Test
    DemoLongOperation,

    // Dashboard
    GetDashboardStats,

    // Mailbox operations
    GetMailboxes,
    GetMailboxDetails,
    GetMailboxPermissions,
    SetMailboxPermission,
    ApplyPermissionsDeltaPlan,
    SetMailboxFeature,
    UpdateMailboxSettings,
    SetMailboxAutoReplyConfiguration,
    ConvertMailboxToShared,

    // Distribution List operations
    GetDistributionLists,
    GetDistributionListDetails,
    GetGroupMembers,
    ModifyGroupMember,
    PreviewDynamicGroupMembers,
    SetDistributionListSettings
}
