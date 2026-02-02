namespace ExchangeAdmin.Contracts.Messages;

public enum OperationType
{
    ConnectExchangeInteractive,
    DisconnectExchange,
    GetConnectionStatus,

    DetectCapabilities,

    DemoLongOperation,

    GetDashboardStats,

    GetMailboxes,
    GetDeletedMailboxes,
    GetMailboxDetails,
    GetRetentionPolicies,
    SetRetentionPolicy,
    GetMailboxPermissions,
    SetMailboxPermission,
    ApplyPermissionsDeltaPlan,
    SetMailboxFeature,
    UpdateMailboxSettings,
    SetMailboxAutoReplyConfiguration,
    ConvertMailboxToShared,
    ConvertMailboxToRegular,
    RestoreMailbox,
    GetMailboxSpaceReport,

    GetDistributionLists,
    GetDistributionListDetails,
    GetGroupMembers,
    ModifyGroupMember,
    PreviewDynamicGroupMembers,
    SetDistributionListSettings,

    GetMessageTrace,

    GetUserLicenses,
    SetUserLicense,
    GetAvailableLicenses,

    CheckPrerequisites,
    InstallModule
}
