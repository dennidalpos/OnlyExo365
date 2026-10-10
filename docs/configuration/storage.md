# Local storage

| Data | Default path | Owner |
| --- | --- | --- |
| UI, Worker, Supervisor JSON-line logs | `%LocalAppData%\OnlyExo365\logs` | PersistentLogWriter |
| DPAPI password references | `%LocalAppData%\OnlyExo365\ipc-secrets` | ProtectedSecretStore; consumed once, stale files cleaned after 6 hours |
| Excel/CSV reports | `%LocalAppData%\OnlyExo365\exports` | Export services; `ONLYEXO365_EXPORT_DIR` overrides |
| SKU JSON and metadata | `%LocalAppData%\OnlyExo365\LicenseCatalog` | LicenseCatalogFileStore; configured path is used literally |
| Selected UI language | `%AppData%\OnlyExo365\OnlyExo365\preferences.json` | UserPreferencesService |
| Machine configuration | `%ProgramData%\OnlyExo365\OnlyExo365\appsettings.json` | Operator policy |

`ONLYEXO365_LOG_RETENTION_DAYS` accepts a positive integer; default is **14 days**. [PersistentLogWriter](../../src/OnlyExo365.Contracts/Diagnostics/PersistentLogWriter.cs) / [PersistentLogStore](../../src/OnlyExo365.Contracts/Diagnostics/PersistentLogStore.cs) apply retention and export diagnostics. [UserPreferencesService](../../src/OnlyExo365.Shell/Services/UserPreferencesService.cs) uses roaming AppData, so not all state is under LocalAppData.

Uninstall removes installed payload plus the uninstalling account's default logs/IPC secrets. It preserves exports, catalog/preferences and unknown installation files; it does not scan every Windows user's profile or custom export directories. See [packaging](../releases/packaging.md).

Repository clean targets only generated outputs and preserves reports, inputs and personal state. Installer smoke temporarily backs up default product data and retains recovery evidence on failure; use an isolated host. See [workflow](../development/workflow.md) and [qualification](../development/qualification.md).
