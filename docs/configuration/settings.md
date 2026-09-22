# Configuration and Settings

OnlyExo365 applies configuration in a strict 3-tier precedence order:

1. **Application Defaults**: `appsettings.json` in the application binary directory.
2. **Per-Machine Policy**: `%ProgramData%\OnlyExo365\OnlyExo365\appsettings.json`.
3. **Environment Overrides**: `ONLYEXO365_*` system or session environment variables.

Legacy compatibility paths are intentionally unsupported.

---

## Environment Variables Reference

### Authentication & Exchange Online

| Variable | Description | Supported Values |
| :--- | :--- | :--- |
| `ONLYEXO365_AUTH_MODE` | Authentication mechanism | `Interactive`, `DeviceCode`, `AppCertificate`, `ManagedIdentity` |
| `ONLYEXO365_EXO_ENV` | Exchange cloud environment | `O365Default`, `O365USGovGCCHigh`, `O365USGovDoD`, `O365China` |
| `ONLYEXO365_EXO_ORGANIZATION` | Target tenant identifier | Tenant DNS domain (e.g. `contoso.onmicrosoft.com`) or GUID |
| `ONLYEXO365_EXO_DELEGATED_ORGANIZATION` | Delegated partner tenant | Tenant DNS domain or GUID |
| `ONLYEXO365_EXO_UPN_HINT` | Login account hint | Email address / UPN |
| `ONLYEXO365_APP_ID` | Entra App Registration ID | Client Application GUID |
| `ONLYEXO365_CERT_THUMBPRINT` | App certificate thumbprint | SHA1 hex thumbprint (LocalMachine/CurrentUser store) |
| `ONLYEXO365_CERT_SUBJECT` | App certificate subject name | Certificate CN / Subject string |
| `ONLYEXO365_MANAGED_IDENTITY_ACCOUNT_ID` | User-assigned managed identity | Resource ID or client ID |
| `ONLYEXO365_DISABLE_EXO` | Disables Exchange connection | `1` or `true` |

### Microsoft Graph Settings

| Variable | Description | Default / Format |
| :--- | :--- | :--- |
| `ONLYEXO365_ENABLE_GRAPH` | Connect Graph alongside Exchange | `1` (true) or `0` (false) |
| `ONLYEXO365_GRAPH_TENANT_ID` | Tenant GUID for Graph authentication | Entra Tenant ID |
| `ONLYEXO365_GRAPH_SCOPES` | Read scopes for directory data | Semicolon-separated permission names |
| `ONLYEXO365_GRAPH_LICENSE_WRITE_SCOPES` | Scopes for license assignments | Semicolon-separated (e.g. `User.ReadWrite.All`) |
| `ONLYEXO365_DEFAULT_USAGE_LOCATION` | Fallback user license country | Two-letter ISO country code (e.g. `US`, `IT`) |

### Storage, Logs & System

| Variable | Description | Default Path |
| :--- | :--- | :--- |
| `ONLYEXO365_EXPORT_DIR` | Directory for Excel/CSV exports | `%LocalAppData%\OnlyExo365\exports` |
| `ONLYEXO365_LOG_RETENTION_DAYS` | Log cleanup threshold | `30` days |

---

## Runtime Data Locations

All local user state resides under `%LocalAppData%\OnlyExo365\`:

- **Logs**: `%LocalAppData%\OnlyExo365\logs\` (`ui-*.log`, `worker-*.log`, `supervisor-*.log`)
- **IPC Secrets**: `%LocalAppData%\OnlyExo365\ipc-secrets\` (DPAPI-protected session token)
- **Exports**: `%LocalAppData%\OnlyExo365\exports\` (Generated reports and spreadsheets)
- **License Catalog**: `%LocalAppData%\OnlyExo365\LicenseCatalog\` (Local JSON cache and metadata)

---

## License Catalog Configuration

The `licensingCatalog` section in `appsettings.json` controls SKU catalog updates:

```json
{
  "licensingCatalog": {
    "autoUpdateMode": "Weekly",
    "checkOnStartup": true,
    "remoteSource": "https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference",
    "downloadTimeoutSeconds": 30,
    "localCachePath": "%LocalAppData%\\OnlyExo365\\LicenseCatalog"
  }
}
```

- `autoUpdateMode`: `Disabled`, `Daily`, `Weekly`, `Monthly`.
- `checkOnStartup`: When `true`, validates catalog currency during application startup.
