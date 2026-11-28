# Exchange and Graph configuration

Source: [ExchangeConfigurationLoader](../../src/OnlyExo365.Shell/Services/ExchangeConfigurationLoader.cs) and [ExchangeOnlineConfiguration](../../src/OnlyExo365.Contracts/ExchangeOnlineConfiguration.cs).

Resolution order: built-in defaults → binary-directory `appsettings.json` → `%ProgramData%\OnlyExo365\OnlyExo365\appsettings.json` → environment overrides. Only properties present in an Exchange overlay replace preceding values. Legacy compatibility paths are unsupported. Invalid Exchange configuration fails with source-aware diagnostics; catalog loading has a separate fallback policy.

The shipped [appsettings.json](../../src/OnlyExo365.Shell/appsettings.json) selects Interactive, `O365Default`, Graph enabled, four directory read scopes and `LicenseAssignment.ReadWrite.All` for license writes. Built-in contract defaults have no write scopes; do not assume the shipped file disables writes. RBAC/Graph consent and capability checks still apply.

## Environment overrides

| Variable | JSON property / meaning |
| --- | --- |
| `ONLYEXO365_AUTH_MODE` | `authenticationMode`: Interactive, DeviceCode, AppCertificate, ManagedIdentity |
| `ONLYEXO365_EXO_ENV` | `exchangeEnvironmentName`: O365Default, O365USGovGCCHigh, O365USGovDoD, O365China |
| `ONLYEXO365_EXO_ORGANIZATION` | `exchangeOrganization`: tenant domain or GUID |
| `ONLYEXO365_EXO_DELEGATED_ORGANIZATION` | `delegatedOrganization`: partner tenant domain or GUID |
| `ONLYEXO365_EXO_UPN_HINT` | `userPrincipalNameHint`: sign-in UPN |
| `ONLYEXO365_APP_ID` | `applicationId`: client GUID |
| `ONLYEXO365_CERT_THUMBPRINT` | `certificateThumbprint`: certificate selector |
| `ONLYEXO365_CERT_SUBJECT` | `certificateSubjectName`: alternative certificate selector |
| `ONLYEXO365_MANAGED_IDENTITY_ACCOUNT_ID` | `managedIdentityAccountId`: host-specific identity selector; Graph passes it as ClientId |
| `ONLYEXO365_GRAPH_TENANT_ID` | `graphTenantId`: tenant GUID; required for Graph AppCertificate |
| `ONLYEXO365_GRAPH_SCOPES` | `graphScopes`: semicolon-separated directory read permissions |
| `ONLYEXO365_GRAPH_LICENSE_WRITE_SCOPES` | `graphLicenseWriteScopes`: semicolon-separated license write permissions |
| `ONLYEXO365_DEFAULT_USAGE_LOCATION` | `defaultUsageLocation`: two-letter country code |
| `ONLYEXO365_ENABLE_GRAPH` | `enableGraphAfterExchangeConnect`: 1/0 or true/false |

`ONLYEXO365_DISABLE_EXO=1` or `true` blocks initiating Exchange connections from the Shell UI independently of the JSON hierarchy. The IPC session environment variable is internal; see [IPC](../architecture/ipc.md).

Catalog settings belong to [catalog](../licensing/catalog.md); file paths and log/export overrides to [storage](storage.md); connection limitations to [authentication](../operations/authentication.md).
