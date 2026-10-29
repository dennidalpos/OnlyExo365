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

Cloud routing uses one [shared map](../../src/OnlyExo365.Contracts/Data/CloudEnvironments.json), embedded in Contracts and read by the tenant validation helper. Worker, interactive Shell bootstrap and tenant validation select the same Graph environment and Compliance connection/authorization endpoints:

| Exchange environment | Graph environment | Compliance host | Authorization host |
| --- | --- | --- | --- |
| O365Default | Global | ps.compliance.protection.outlook.com | login.microsoftonline.com |
| O365USGovGCCHigh | USGov | ps.compliance.protection.office365.us | login.microsoftonline.us |
| O365USGovDoD | USGovDoD | l5.ps.compliance.protection.office365.us | login.microsoftonline.us |
| O365China | China | ps.compliance.protection.partner.outlook.cn | login.chinacloudapi.cn |

Unknown clouds, including the legacy O365GermanyCloud value, fail explicitly in Graph/Compliance routing. Delegated Graph in China requires a locally registered ApplicationId; delegated connections use the configured ClientId when supplied. The four Graph modes and supported Compliance modes have offline argument-binding coverage across the four clouds. The installed EXO 3.10.1 command signature was inspected without connecting: it exposes neither Device nor ManagedIdentity for Connect-IPPSSession. [Microsoft Graph authentication](https://learn.microsoft.com/en-us/powershell/microsoftgraph/authentication-commands?view=graph-powershell-1.0) and [Connect-IPPSSession](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/connect-ippssession?view=exchange-ps) support these choices; real cloud authentication remains unqualified.

Compliance app-certificate authentication resolves an exact CertificateSubjectName in CurrentUser/My to exactly one currently valid certificate with a private key and passes its thumbprint to Connect-IPPSSession. Missing or ambiguous usable certificates fail before connecting. An explicit thumbprint takes precedence. Worker and tenant validation share these selection rules; offline fixture tests do not qualify real certificate authentication. Interactive sign-in has a five-minute budget; cancellation terminates the owned authentication process tree.

Compliance uses `Connect-IPPSSession` separately from Exchange. eDiscovery uses a separate `-EnableSearchOnlySession` connection. Microsoft documents delegated search requirements and states app-only eDiscovery is unsupported; its best-effort legacy guidance does not qualify this application. See [Security & Compliance connection guidance](https://learn.microsoft.com/en-us/powershell/exchange/connect-to-scc-powershell?view=exchange-ps). Real module/authentication qualification is unverified in 02-areas/products/OnlyExo365/tracker.json (nel dossier SecondBrain) (SecondBrain: `02-areas/products/OnlyExo365/tracker.json`).
