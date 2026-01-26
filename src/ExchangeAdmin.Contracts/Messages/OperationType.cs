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
    GetMailboxSpaceReport,

    
    GetDistributionLists,
    GetDistributionListDetails,
    GetGroupMembers,
    ModifyGroupMember,
    PreviewDynamicGroupMembers,
    SetDistributionListSettings
}
