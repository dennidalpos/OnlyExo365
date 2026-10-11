# Modules and authentication

Tools installs exact approved versions side by side. Newer installed versions do not replace them. Discovery/import uses the shared [module policy](../../src/OnlyExo365.Contracts/Data/PowerShellModuleBootstrapPolicy.json); the current version table is in [workflow](../development/workflow.md). The Graph bundle comprises Authentication, Users, Users.Actions and Identity.DirectoryManagement; no additional duplicate Graph installation step is needed.

External interactive bootstrap requires PowerShell 7.6+ in PATH. Worker embeds its own PowerShell SDK. Missing approved modules or failed import block the affected connection with an explicit error while keeping Worker/Tools available. Reinstall the approved version and reconnect. Module installation needs Gallery connectivity and permission to write the selected module scope.

| Mode | Connection constraints |
| --- | --- |
| Interactive | Delegated browser sign-in; Compliance/eDiscovery needs an interactive host |
| DeviceCode | Supported for Exchange/Graph; unsupported for the application's Compliance connection |
| AppCertificate | Requires configured app/certificate; Graph also requires GraphTenantId and application consent; eDiscovery is rejected |
| ManagedIdentity | Requires a supported identity host; unsupported for the application's Compliance connection |

Permissions are operation-specific, checked through capability discovery and least-privilege guidance; one Exchange role does not grant every feature. License writes require the configured Graph write scope and consent. See [settings](../configuration/settings.md).

Compliance uses `Connect-IPPSSession` separately from Exchange. eDiscovery uses a separate `-EnableSearchOnlySession` connection. Microsoft documents delegated search requirements and states app-only eDiscovery is unsupported; its best-effort legacy guidance does not qualify this application. See [Security & Compliance connection guidance](https://learn.microsoft.com/en-us/powershell/exchange/connect-to-scc-powershell?view=exchange-ps). Real module/authentication qualification is unverified in [PROJECT_STATUS.json](../../PROJECT_STATUS.json).
