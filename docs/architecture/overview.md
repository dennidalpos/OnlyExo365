# Architecture Overview

This document describes the repository structure and runtime model that are verifiable from source, project files, scripts, installer authoring, and CI configuration.

## Runtime Model

OnlyExo365 is a Windows-first desktop application for Exchange Online operations.

The runtime is split into:

- `ExchangeAdmin.Presentation.exe`: WPF operator shell.
- `ExchangeAdmin.Worker.exe`: separate worker host for PowerShell-backed operations.
- `ExchangeAdmin.Contracts`: DTOs, IPC contracts, configuration contracts, and diagnostics shared across processes.
- Named-pipe IPC plus a DPAPI-protected session token passed through `EXCHANGEADMIN_IPC_SESSION_TOKEN`.

The presentation process owns the operator UI, local configuration loading, runtime catalog cache, and worker supervision. The worker process owns command dispatching and Exchange Online/Microsoft Graph PowerShell integration.

## Source Layout

- `src/ExchangeAdmin.Presentation`: WPF app, views, view models, UI services, app configuration, localization, assets, and catalog update service.
- `src/ExchangeAdmin.Worker`: worker host, operation dispatching, PowerShell integration, and embedded data resources.
- `src/ExchangeAdmin.Contracts`: shared contracts, IPC constants, Exchange configuration contracts, diagnostics, and DTOs.
- `src/ExchangeAdmin.Application`: application services and use cases.
- `src/ExchangeAdmin.Infrastructure`: IPC client and infrastructure adapters.
- `src/ExchangeAdmin.Domain`: domain types, errors, and resilience primitives.
- `tests/ExchangeAdmin.Tests`: xUnit coverage for application, worker, IPC, configuration, packaging scripts, localization, and WPF-adjacent behavior.
- `scripts/`: canonical repository entrypoints for bootstrap, compile, test, packaging, publish, cleanup, and catalog refresh.
- `build/`: release, packaging, smoke, signing, security, coverage, architecture, and validation helpers.
- `installer/`: Inno Setup authoring for `OnlyExo365.Setup.exe`.
- `.github/`: Windows release baseline action and release validation workflows.

## Toolchain Baseline

The repository is Windows-first.

Required development baseline:

- Windows
- PowerShell 7 for repository scripts
- .NET SDK `10.0.202`, pinned in `global.json` with `rollForward` disabled
- NuGet locked restore through versioned `packages.lock.json` files
- Inno Setup 6 when packaging the setup EXE locally

Current target frameworks:

- `ExchangeAdmin.Presentation`: `net10.0-windows`
- `ExchangeAdmin.Worker`: `net10.0`
- `ExchangeAdmin.Tests`: `net10.0-windows`
- `ExchangeAdmin.Application`, `ExchangeAdmin.Contracts`, `ExchangeAdmin.Domain`, `ExchangeAdmin.Infrastructure`: `net8.0`

Notable package baselines:

- `Microsoft.PowerShell.SDK` `7.6.0` in the worker
- `System.Security.Cryptography.Xml` `10.0.6` in the worker
- `CommunityToolkit.Mvvm` `8.2.2` and `DocumentFormat.OpenXml` `3.2.0` in the presentation project

The PowerShell SDK and .NET baseline decision is documented in [../decisions/2026-04-11-powershell-sdk-major-upgrade-strategy.md](../decisions/2026-04-11-powershell-sdk-major-upgrade-strategy.md).

## Runtime Configuration Ownership

The versioned default configuration lives in `src/ExchangeAdmin.Presentation/appsettings.json`.

Operator-facing configuration, environment variable overrides, runtime data paths, authentication modes, and first-run guidance are documented in [../operations/operator-setup.md](../operations/operator-setup.md).

## Build and Release Ownership

Build, test, and CI gate commands are documented in [../development/build-test.md](../development/build-test.md).

Packaging, smoke validation, signing, release assets, and tenant validation are documented in [../releases/verification.md](../releases/verification.md).

## Microsoft 365 SKU Catalog

The worker embeds `src/ExchangeAdmin.Worker/Data/Microsoft365SkuCatalog.json` as a fallback catalog. The presentation process can also maintain a runtime catalog cache.

Manual refresh of the versioned embedded catalog is documented in [../maintenance/licensing-catalog.md](../maintenance/licensing-catalog.md).
