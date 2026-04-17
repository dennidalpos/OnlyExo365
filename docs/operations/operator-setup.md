# Operator Setup and Runtime Configuration

This guide covers packaged-app installation, first-run checks, and runtime configuration. Source builds and test gates are covered in [../development/build-test.md](../development/build-test.md). Packaging, signing, and release validation are covered in [../releases/verification.md](../releases/verification.md).

## Minimum Requirements

Target workstation:

- Windows 10 or Windows 11
- .NET Desktop Runtime 10 for the framework-dependent payload
- internet access to Microsoft 365 endpoints
- an account or app identity with permissions for the intended Exchange Online and Graph operations

The `Tools` page verifies the operator prerequisites implemented by the current app:

- PowerShell 7
- supported current-user execution policy: `RemoteSigned`, `Unrestricted`, or `Bypass`
- approved `ExchangeOnlineManagement` module version
- approved Microsoft Graph module bundle

## Installation

Use one of the repository-produced artifacts:

- `artifacts/packages/OnlyExo365.Setup.exe`
- `artifacts/publish/win-x64/ExchangeAdmin.Presentation.exe` on 64-bit Windows
- `artifacts/publish/win-x86/ExchangeAdmin.Presentation.exe` on 32-bit Windows

The setup EXE is authored as a per-machine Inno Setup installer:

- default install directory: `C:\Program Files\OnlyExo365`
- requires administrator privileges
- creates Start Menu and Desktop shortcuts
- registers uninstall information
- removes the installed app directory, logs, IPC secrets, and default export directory during uninstall

For direct publish-folder evaluation, launch `ExchangeAdmin.Presentation.exe` from the runtime-specific publish directory that matches the target host architecture.

## First Run

1. Start the app.
2. Open `Tools`.
3. Verify PowerShell 7, execution policy, Exchange Online Management, and the Microsoft Graph bundle.
4. Start the worker.
5. Connect to Exchange Online.
6. Confirm the shell reports coherent worker, Exchange, and Graph state.

## Configuration Sources

The versioned default configuration is `src/ExchangeAdmin.Presentation/appsettings.json`.

Runtime precedence:

1. `appsettings.json` in the application directory
2. `%ProgramData%\OnlyExo365\appsettings.json`
3. `EXCHANGEADMIN_*` environment variable overrides

For upgrade compatibility, the app still reads `%ProgramData%\OnlyExo365\ExchangeAdmin\appsettings.json` when the new shared path is absent.

For shared workstation configuration, prefer:

```powershell
New-Item -ItemType Directory -Force "$env:ProgramData\OnlyExo365"
```

Then place `appsettings.json` in that directory.

## Minimal Shared Configuration

```json
{
  "ExchangeOnline": {
    "authenticationMode": "Interactive",
    "exchangeEnvironmentName": "O365Default",
    "exchangeOrganization": "contoso.onmicrosoft.com",
    "graphTenantId": "contoso.onmicrosoft.com",
    "defaultUsageLocation": "IT",
    "graphLicenseWriteScopes": [
      "LicenseAssignment.ReadWrite.All"
    ],
    "enableGraphAfterExchangeConnect": true
  }
}
```

Real constraints:

- `exchangeOrganization` and `graphTenantId` accept a tenant domain or tenant GUID.
- `defaultUsageLocation`, when set, must be a two-letter country or region code.
- `graphScopes` and `graphLicenseWriteScopes` use `;`-separated values when supplied by environment variable.
- `enableGraphAfterExchangeConnect` controls whether Graph connects after Exchange connects.

## Authentication Modes

Supported values:

- `Interactive`
- `DeviceCode`
- `AppCertificate`
- `ManagedIdentity`

`AppCertificate` requires:

- `applicationId`
- `exchangeOrganization`
- `certificateThumbprint` or `certificateSubjectName`
- `graphTenantId` when Graph is enabled after Exchange connects

`ManagedIdentity` requires:

- `exchangeOrganization`

Use `Interactive` unless a concrete unattended or managed-identity scenario is required.

## Environment Overrides

Supported Exchange and Graph overrides:

- `EXCHANGEADMIN_EXO_ENV`
- `EXCHANGEADMIN_AUTH_MODE`
- `EXCHANGEADMIN_EXO_ORGANIZATION`
- `EXCHANGEADMIN_EXO_DELEGATED_ORGANIZATION`
- `EXCHANGEADMIN_EXO_UPN_HINT`
- `EXCHANGEADMIN_APP_ID`
- `EXCHANGEADMIN_CERT_THUMBPRINT`
- `EXCHANGEADMIN_CERT_SUBJECT`
- `EXCHANGEADMIN_MANAGED_IDENTITY_ACCOUNT_ID`
- `EXCHANGEADMIN_GRAPH_TENANT_ID`
- `EXCHANGEADMIN_GRAPH_SCOPES`
- `EXCHANGEADMIN_GRAPH_LICENSE_WRITE_SCOPES`
- `EXCHANGEADMIN_DEFAULT_USAGE_LOCATION`
- `EXCHANGEADMIN_ENABLE_GRAPH`

Additional local runtime overrides:

- `EXCHANGEADMIN_EXPORT_DIR`: overrides `%LocalAppData%\OnlyExo365\exports`
- `EXCHANGEADMIN_LOG_RETENTION_DAYS`: overrides the default 14-day persistent log retention when set to a positive integer

## Runtime Data

Default local paths:

- logs: `%LocalAppData%\OnlyExo365\logs`
- DPAPI IPC secrets: `%LocalAppData%\OnlyExo365\ipc-secrets`
- exports: `%LocalAppData%\OnlyExo365\exports`
- runtime Microsoft 365 SKU catalog cache: `%LocalAppData%\OnlyExo365\LicenseCatalog`

The setup EXE uninstall removes the installed app directory plus the default logs, IPC secret, and export directories.

## Microsoft 365 SKU Catalog Updates

The app has a versioned embedded catalog and a presentation-side runtime cache.

The default `licensingCatalog` configuration:

- checks on startup
- uses `Daily` auto-update mode
- downloads from Microsoft Learn
- stores the runtime cache under `%LocalAppData%\OnlyExo365\LicenseCatalog` unless `localCachePath` is set

Set `autoUpdateMode` to `Disabled` in configuration if runtime catalog downloads are not allowed in an environment.

## Troubleshooting

- If the app does not start, verify that .NET Desktop Runtime 10 is installed.
- If the setup EXE fails under `Program Files`, rerun it with administrator privileges.
- If `Tools` reports an unsupported execution policy, run `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned` in PowerShell 7.
- If `Tools` reports missing PowerShell, install PowerShell 7 and relaunch the app.
- If Exchange Online or Graph modules are missing or in drift, use the actions exposed by `Tools`.
- If configuration is not applied, re-check the source precedence above.
