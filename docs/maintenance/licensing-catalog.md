# Microsoft 365 SKU Catalog Maintenance

This document covers the versioned embedded catalog at `src/OnlyExo365.Worker/Data/Microsoft365SkuCatalog.json`.

The app also has a runtime cache managed by the presentation process. Runtime cache behavior is configured by the `licensingCatalog` section in `src/OnlyExo365.Shell/appsettings.json` and described in [../operations/operator-setup.md](../operations/operator-setup.md).

## Purpose

The embedded catalog is the fallback catalog used by the worker and presentation resolver when no valid local runtime cache exists.

It supports:

- friendly SKU and service-plan names
- licensing DTO enrichment
- Exchange Online plan recognition inside Microsoft 365 licenses

## Canonical Refresh Command

```powershell
pwsh ./scripts/agents/refresh-microsoft365-sku-catalog.ps1
```

The script:

- resolves the current Microsoft Learn licensing CSV link
- downloads the CSV into `artifacts/tmp/licensing-catalog`
- regenerates `src/OnlyExo365.Worker/Data/Microsoft365SkuCatalog.json`
- updates the `generatedOn` field

## Expected JSON Shape

The generated file contains:

- `generatedOn`
- `source`
- `csvDownload`
- `entries`

Entries are normalized and ordered by `skuPartNumber` and `skuId`.
Each entry includes `servicePlans` with `servicePlanName`, `servicePlanId`, and `friendlyName` when the source CSV provides plan rows.

## Current Versioned Snapshot

The current versioned catalog reports:

- `generatedOn = 2026-04-11`
- `entries = 617`

## Minimum Verification After Refresh

```powershell
pwsh ./scripts/agents/compile.ps1 -Configuration Debug -RuntimeIdentifier win-x64
pwsh ./scripts/agents/test.ps1 -Configuration Debug -RuntimeIdentifier win-x64
```

