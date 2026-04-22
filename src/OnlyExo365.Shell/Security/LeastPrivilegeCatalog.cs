using OnlyExo365.Contracts;

namespace OnlyExo365.Shell.Security;

public static class LeastPrivilegeCatalog
{
    private static readonly IReadOnlyList<ExchangeAuthenticationMode> AllAuthenticationModes =
    [
        ExchangeAuthenticationMode.Interactive,
        ExchangeAuthenticationMode.DeviceCode,
        ExchangeAuthenticationMode.AppCertificate,
        ExchangeAuthenticationMode.ManagedIdentity
    ];

    public const string MobileDevicesInventory = "mobile-devices.inventory";
    public const string MigrationBatches = "migration.batches";
    public const string PermissionsRoleGroups = "permissions.role-groups";
    public const string MessageTraceRead = "message-trace.read";
    public const string MailFlowTransportAndRouting = "mail-flow.transport-and-routing";
    public const string ComplianceAuditAndEDiscovery = "compliance.audit-and-ediscovery";
    public const string MailSecurityBaseline = "mail-security.baseline";
    public const string MailboxLicensingRead = "mailboxes.licensing-read";
    public const string MailboxLicensingWrite = "mailboxes.licensing-write";

    public static IReadOnlyList<LeastPrivilegeFeatureDefinition> All { get; } =
    [
        new LeastPrivilegeFeatureDefinition
        {
            FeatureId = MobileDevicesInventory,
            ModuleName = "Mobile Devices",
            FeatureName = "ActiveSync inventory and actions",
            Description = "Inventory device, policy load, allow/block/quarantine and remote wipe.",
            AllowedAuthenticationModes = AllAuthenticationModes,
            RequiredCmdletsAll =
            [
                "Get-MobileDevice",
                "Get-MobileDeviceStatistics",
                "Get-MobileDeviceMailboxPolicy",
                "Get-CASMailbox",
                "Set-CASMailbox",
                "Clear-MobileDevice"
            ],
            RecommendedExchangeRoles =
            [
                "Recipient Management"
            ],
            Dependencies =
            [
                "Exchange Online session"
            ],
            Notes = "Use a custom RBAC role group narrowed to mobile device and CAS mailbox cmdlets when Recipient Management is broader than required."
        },
        new LeastPrivilegeFeatureDefinition
        {
            FeatureId = MigrationBatches,
            ModuleName = "Migration",
            FeatureName = "Batch inventory, endpoint test and batch lifecycle",
            Description = "Read migration batches, manage endpoints, run preflight and start/complete/remove batches.",
            AllowedAuthenticationModes = AllAuthenticationModes,
            RequiredCmdletsAll =
            [
                "Get-MigrationBatch",
                "Get-MigrationEndpoint",
                "New-MigrationEndpoint",
                "Set-MigrationEndpoint",
                "Test-MigrationServerAvailability",
                "New-MigrationBatch",
                "Start-MigrationBatch",
                "Complete-MigrationBatch",
                "Remove-MigrationBatch"
            ],
            RecommendedExchangeRoles =
            [
                "Organization Management"
            ],
            Dependencies =
            [
                "Exchange Online session"
            ],
            Notes = "Prefer a custom Exchange RBAC role group containing only migration roles/cmdlets; Organization Management remains the broad built-in fallback."
        },
        new LeastPrivilegeFeatureDefinition
        {
            FeatureId = PermissionsRoleGroups,
            ModuleName = "Permissions",
            FeatureName = "RBAC role groups and membership",
            Description = "Read role groups, inspect scopes and manage group membership.",
            AllowedAuthenticationModes = AllAuthenticationModes,
            RequiredCmdletsAll =
            [
                "Get-RoleGroup",
                "Get-RoleGroupMember",
                "New-RoleGroup",
                "Add-RoleGroupMember",
                "Remove-RoleGroupMember"
            ],
            RecommendedExchangeRoles =
            [
                "Role Management",
                "Organization Management"
            ],
            Dependencies =
            [
                "Exchange Online session"
            ],
            Notes = "Prefer a custom role group with the Role Management management role instead of assigning Organization Management."
        },
        new LeastPrivilegeFeatureDefinition
        {
            FeatureId = MessageTraceRead,
            ModuleName = "Message Trace",
            FeatureName = "Trace search and detail",
            Description = "Run message trace and inspect per-recipient detail.",
            AllowedAuthenticationModes = AllAuthenticationModes,
            RequiredCmdletsAny =
            [
                "Get-MessageTraceV2",
                "Get-MessageTrace"
            ],
            RecommendedExchangeRoles =
            [
                "Hygiene Management"
            ],
            Dependencies =
            [
                "Exchange Online Protection telemetry"
            ],
            Notes = "Get-MessageTraceDetailV2 or Get-MessageTraceDetail remains optional but is required for the detail drawer."
        },
        new LeastPrivilegeFeatureDefinition
        {
            FeatureId = MailFlowTransportAndRouting,
            ModuleName = "Mail Flow",
            FeatureName = "Transport rules, connectors and routing domains",
            Description = "Manage transport rules, inbound or outbound connectors, accepted domains, remote domains and organization relationships.",
            AllowedAuthenticationModes = AllAuthenticationModes,
            RequiredCmdletsAny =
            [
                "Get-TransportRule",
                "Get-InboundConnector",
                "Get-OutboundConnector",
                "Get-AcceptedDomain",
                "Get-RemoteDomain",
                "Get-OrganizationRelationship"
            ],
            RecommendedExchangeRoles =
            [
                "Hygiene Management",
                "Organization Management"
            ],
            Dependencies =
            [
                "Exchange transport configuration"
            ],
            Notes = "Least privilege usually means custom routing or hygiene role groups instead of broad Organization Management."
        },
        new LeastPrivilegeFeatureDefinition
        {
            FeatureId = ComplianceAuditAndEDiscovery,
            ModuleName = "Compliance",
            FeatureName = "Audit log, compliance search, purge and hold",
            Description = "Open Purview session during the initial connect and run audit/eDiscovery workflows, with explicit retry only when the session decays or cmdlets are missing.",
            AllowedAuthenticationModes = AllAuthenticationModes,
            RequiredCmdletsAll =
            [
                "Search-UnifiedAuditLog",
                "Get-ComplianceSearch",
                "New-ComplianceSearch",
                "Start-ComplianceSearch",
                "Remove-ComplianceSearch",
                "New-ComplianceSearchAction",
                "New-CaseHoldPolicy",
                "New-CaseHoldRule"
            ],
            RecommendedPurviewRoles =
            [
                "eDiscovery Manager",
                "eDiscovery Administrator",
                "Audit Logs"
            ],
            Dependencies =
            [
                "Purview / Security & Compliance PowerShell",
                "Connect-IPPSSession during the initial connect flow"
            ],
            RequiresAdditionalSessionValidation = true,
            Notes = "Compliance cmdlets are validated after the initial Connect-IPPSSession bootstrap and retried explicitly only if the session decays or the requested cmdlet is missing. Exchange RBAC alone is not sufficient for this module."
        },
        new LeastPrivilegeFeatureDefinition
        {
            FeatureId = MailSecurityBaseline,
            ModuleName = "Mail Security",
            FeatureName = "DKIM, anti-spam, anti-phish, malware, quarantine and outbound spam",
            Description = "Load and update EOP / Defender policy surfaces exposed by the worker.",
            AllowedAuthenticationModes = AllAuthenticationModes,
            RequiredCmdletsAny =
            [
                "Get-DkimSigningConfig",
                "Get-HostedContentFilterPolicy",
                "Get-AntiPhishPolicy",
                "Get-MalwareFilterPolicy",
                "Get-QuarantinePolicy",
                "Get-HostedOutboundSpamFilterPolicy"
            ],
            RecommendedExchangeRoles =
            [
                "Hygiene Management"
            ],
            RecommendedDefenderRoles =
            [
                "Security Administrator",
                "Quarantine Administrator"
            ],
            Dependencies =
            [
                "Exchange Online Protection / Defender for Office 365"
            ],
            Notes = "No single built-in role is minimal for every sub-surface. Quarantine-only operations can stay on Quarantine Administrator while broader policy editing usually needs Hygiene Management or Security Administrator."
        },
        new LeastPrivilegeFeatureDefinition
        {
            FeatureId = MailboxLicensingRead,
            ModuleName = "Dashboard / Mailboxes",
            FeatureName = "Graph-backed insights and license inventory",
            Description = "Dashboard license inventory, admin role insights and mailbox license detail.",
            AllowedAuthenticationModes = AllAuthenticationModes,
            RequiredGraphScopes =
            [
                "Organization.Read.All",
                "Directory.Read.All",
                "RoleManagement.Read.Directory",
                "User.Read.All"
            ],
            Dependencies =
            [
                "Microsoft Graph initial session"
            ],
            Notes = "Read-only Graph scopes stay in GraphScopes so the dashboard and mailbox detail can bootstrap during the initial connect without enabling license assignment changes."
        },
        new LeastPrivilegeFeatureDefinition
        {
            FeatureId = MailboxLicensingWrite,
            ModuleName = "Mailboxes",
            FeatureName = "Mailbox licensing write actions",
            Description = "Assign and remove licenses on mailbox-backed users.",
            AllowedAuthenticationModes = AllAuthenticationModes,
            RequiredGraphScopes =
            [
                "Organization.Read.All",
                "Directory.Read.All",
                "RoleManagement.Read.Directory",
                "User.Read.All",
                "LicenseAssignment.ReadWrite.All"
            ],
            Dependencies =
            [
                "Microsoft Graph initial session"
            ],
            Notes = "License write scopes are opt-in via graphLicenseWriteScopes / ONLYEXO365_GRAPH_LICENSE_WRITE_SCOPES. AppCertificate mode also requires GraphTenantId to avoid implicit tenant discovery."
        }
    ];

}

