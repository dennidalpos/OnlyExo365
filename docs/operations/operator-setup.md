# Operator Setup and Runtime Configuration

This guide covers installation, first-run checks, and runtime configuration.

## Minimum Requirements

- Windows 10 or Windows 11 x64
- .NET Desktop Runtime 10
- PowerShell 7
- internet access to Microsoft 365 endpoints
- an Exchange Online / Microsoft Graph identity with the permissions required for the selected workflows

## Installation

Supported outputs:

- `artifacts/packages/OnlyExo365.Setup.exe`
- `artifacts/publish/win-x64/OnlyExo365.Shell.exe`

The setup EXE is a per-machine Inno Setup installer:

- default install directory: `C:\Program Files\OnlyExo365`
- administrator privileges required
- common Start Menu and Desktop shortcuts
- uninstall registration
- uninstall cleanup for install directory, logs, IPC secrets, and default export directory

## First Run

1. Start the app.
2. Open `Tools`.
3. Verify PowerShell 7, execution policy, Exchange Online Management, and Microsoft Graph prerequisites.
4. Start the worker.
5. Connect to Exchange Online.
6. Confirm coherent shell, worker, Exchange, and Graph state.

## Configuration Resolution

Resolution order:

1. `appsettings.json` in the application directory
2. `%ProgramData%\OnlyExo365\OnlyExo365\appsettings.json`
3. `ONLYEXO365_*` environment variables

No legacy compatibility directory is part of the supported baseline.

## Environment Overrides

- `ONLYEXO365_EXO_ENV`
- `ONLYEXO365_AUTH_MODE`
- `ONLYEXO365_EXO_ORGANIZATION`
- `ONLYEXO365_EXO_DELEGATED_ORGANIZATION`
- `ONLYEXO365_EXO_UPN_HINT`
- `ONLYEXO365_APP_ID`
- `ONLYEXO365_CERT_THUMBPRINT`
- `ONLYEXO365_CERT_SUBJECT`
- `ONLYEXO365_MANAGED_IDENTITY_ACCOUNT_ID`
- `ONLYEXO365_GRAPH_TENANT_ID`
- `ONLYEXO365_GRAPH_SCOPES`
- `ONLYEXO365_GRAPH_LICENSE_WRITE_SCOPES`
- `ONLYEXO365_DEFAULT_USAGE_LOCATION`
- `ONLYEXO365_ENABLE_GRAPH`
- `ONLYEXO365_DISABLE_EXO`
- `ONLYEXO365_EXPORT_DIR`
- `ONLYEXO365_LOG_RETENTION_DAYS`

## Runtime Data

Default local paths:

- logs: `%LocalAppData%\OnlyExo365\logs`
- IPC secrets: `%LocalAppData%\OnlyExo365\ipc-secrets`
- exports: `%LocalAppData%\OnlyExo365\exports`
- license catalog cache: `%LocalAppData%\OnlyExo365\LicenseCatalog`
