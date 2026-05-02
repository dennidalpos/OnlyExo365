# Architecture Overview

This repository is consolidated around a single Windows/x64 runtime model.

## Runtime Model

- `OnlyExo365.Shell.exe`: WPF shell, configuration loading, persistent logging, license catalog cache handling, and worker supervision
- `OnlyExo365.Worker.exe`: PowerShell-backed execution host, runspace lifecycle, operation dispatching, and Exchange/Graph command execution
- `OnlyExo365.Contracts`: shared IPC constants, message contracts, DTOs, configuration types, result/error primitives, and diagnostics
- IPC transport: named pipes plus a DPAPI-protected session token exposed through `ONLYEXO365_IPC_SESSION_TOKEN`

The shell owns the desktop process boundary. The worker owns PowerShell execution. No separate `Application`, `Infrastructure`, or `Domain` projects remain in the production architecture.

## Repo-Local Layering Rule

OnlyExo365 intentionally keeps the production architecture consolidated into Shell, Worker, and Contracts. The clean-architecture rule is enforced inside those projects rather than through separate project names:

- domain rules are pure evaluators, request/response decisions, capability rules, and error/result models with no UI, storage, installer, or PowerShell process ownership
- application use cases orchestrate domain rules and worker calls without owning WPF controls, named-pipe transport, files, installers, or command text generation
- infrastructure code owns named pipes, DPAPI-backed secret files, persistent logs, configuration file access, package/install scripts, catalog download/storage, and PowerShell execution
- presentation code owns WPF views, view models, localization display text, commands, dialogs, and navigation state

New code should strengthen these internal boundaries first. Do not add `Domain`, `Application`, or `Infrastructure` projects unless a planned architecture change also moves real responsibilities into them in the same work.

## Source Layout

- `src/OnlyExo365.Shell`
- `src/OnlyExo365.Worker`
- `src/OnlyExo365.Contracts`
- `tests/OnlyExo365.Tests`
- `scripts/`: canonical local entrypoints, agent automation scripts, and shared script helpers
- `build/`: CI, release, security, smoke, signing, and validation helpers
- `installer/`: Inno Setup authoring for `OnlyExo365.Setup.exe`
- `.github/`: Windows release baseline action and release workflows

## Toolchain Baseline

- Windows x64
- PowerShell 7+ (`pwsh`)
- .NET SDK `10.0.203`, pinned in `global.json` with roll-forward disabled
- Inno Setup 6 for local packaging
- solution/project baseline: `net10.0` or `net10.0-windows` only
- supported runtime identifier: `win-x64`

## Configuration Ownership

- versioned defaults: `src/OnlyExo365.Shell/appsettings.json`
- shared machine override: `%ProgramData%\OnlyExo365\OnlyExo365\appsettings.json`
- environment overrides: `ONLYEXO365_*`

Only the current configuration path is supported. Legacy compatibility directories are intentionally not part of the baseline.
