# Microsoft 365 SKU Catalog

OnlyExo365 provides human-readable names and service-plan mappings for Microsoft 365 licenses using a dual-tier catalog system.

## Dual-Tier Architecture

1. **Embedded Worker Catalog**:
   - Located at `src/OnlyExo365.Worker/Data/Microsoft365SkuCatalog.json`.
   - Bundled directly into the worker binary as an embedded resource.
   - Acts as the immutable baseline fallback whenever local caches are missing or corrupted.
2. **Dynamic Presentation Cache**:
   - Managed directly by `OnlyExo365.Shell` via `LicenseCatalogUpdateService`.
   - Stored under `%LocalAppData%\OnlyExo365\LicenseCatalog\`.
   - Periodically checks Microsoft documentation in the background and atomically updates the local cache without requiring worker interaction.

---

## Catalog Refresh Procedure

To regenerate the embedded baseline catalog from official Microsoft Learn data:

```powershell
pwsh ./scripts/agents/refresh-microsoft365-sku-catalog.ps1
```

### Script Actions

1. Queries the Microsoft Learn Licensing and Service Plan Reference page.
2. Downloads the official CSV mapping table.
3. Parses, groups, and normalizes SKU entries and service plan definitions.
4. Generates a formatted JSON snapshot at `src/OnlyExo365.Worker/Data/Microsoft365SkuCatalog.json`.

---

## Verification After Refresh

After updating the catalog snapshot, verify build integrity:

```powershell
pwsh ./scripts/agents/compile.ps1 -Configuration Debug -RuntimeIdentifier win-x64
pwsh ./scripts/agents/test.ps1 -Configuration Debug -RuntimeIdentifier win-x64
```
