# Architecture Overview

This repository is consolidated around a single Windows/x64 runtime model.

## Runtime Model

- `OnlyExo365.Shell.exe`: WPF shell, configuration loading, persistent logging, license catalog cache handling, and worker supervision
- `OnlyExo365.Worker.exe`: PowerShell-backed execution host, runspace lifecycle, operation dispatching, and Exchange/Graph command execution
- `OnlyExo365.Contracts`: shared IPC constants, message contracts, DTOs, configuration types, result/error primitives, and diagnostics
- IPC transport: named pipes plus a DPAPI-protected session token exposed through `ONLYEXO365_IPC_SESSION_TOKEN`

The shell owns the desktop process boundary. The worker owns PowerShell execution. No separate `Application`, `Infrastructure`, or `Domain` projects remain in the production architecture.

## Source Layout

- `src/OnlyExo365.Shell`
- `src/OnlyExo365.Worker`
- `src/OnlyExo365.Contracts`
- `tests/OnlyExo365.Tests`
- `scripts/`: canonical local entrypoints
- `build/`: CI, release, security, smoke, signing, and validation helpers
- `installer/`: Inno Setup authoring for `OnlyExo365.Setup.exe`
- `.github/`: Windows release baseline action and release workflows

## Toolchain Baseline

- Windows
- PowerShell 7
- .NET SDK `10.0.203`, pinned in `global.json`
- Inno Setup 6 for local packaging
- solution/project baseline: `net10.0` or `net10.0-windows` only
- supported runtime identifier: `win-x64`

## Configuration Ownership

- versioned defaults: `src/OnlyExo365.Shell/appsettings.json`
- shared machine override: `%ProgramData%\OnlyExo365\appsettings.json`
- environment overrides: `ONLYEXO365_*`

Only the current configuration path is supported. Legacy compatibility directories are intentionally not part of the baseline.
