# Microsoft 365 SKU catalog

Worker embeds [Microsoft365SkuCatalog.json](../../src/OnlyExo365.Worker/Data/Microsoft365SkuCatalog.json) as its baseline. Shell `LicenseCatalogUpdateService` loads the local catalog or falls back to that embedded resource; background downloads update presentation names without IPC. `PresentationSkuNameResolver` swaps immutable indexes, preferring part number over ID, then a humanized fallback.

## Shell cache settings

`licensingCatalog` is loaded from binary-directory and machine appsettings. Invalid catalog sections retain the preceding/default configuration, unlike invalid Exchange configuration. Persisted metadata update mode overrides the configured mode. No catalog environment overlay exists.

| Property | Default / behavior |
| --- | --- |
| `autoUpdateMode` | Daily; supported: Disabled, Daily, Monthly |
| `checkOnStartup` | true; downloads only when due |
| `remoteSource` | [Microsoft licensing reference](https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference) |
| `downloadTimeoutSeconds` | 30 |
| `localCachePath` | null → [default LocalAppData path](../configuration/storage.md); custom text is used literally, without `%VAR%` expansion |

Daily/Monthly timers use 24 hours/30 days; due checks use 23 hours/29 days. Disabled stops scheduling/startup refresh; manual forced refresh remains available. Catalog writes validate JSON then replace atomically. Shell catalog downloads use the built-in .NET `TextFieldParser` for quoted fields, embedded newlines and malformed-row detection; no new package is required.

Due checks use the previous LastCheckedUtc value. A check that is not due leaves that timestamp unchanged; an actual attempt records its start and re-arms the one-shot timer after success, failure or cancellation. Failed downloads preserve the existing catalog. These behaviors have deterministic clock/HTTP coverage in LicenseCatalogUpdateServiceTests.

## Embedded catalog maintenance

```powershell
pwsh ./scripts/agents/refresh-microsoft365-sku-catalog.ps1
```

The script downloads Microsoft's CSV, validates required columns/GUIDs/product identities, groups SKUs/plans and atomically replaces the embedded JSON. Single-SKU output remains an array. Review generated changes; do not hand-edit the snapshot. For a concrete refresh problem, use the native script regressions and catalog tests described in [workflow](../development/workflow.md).
