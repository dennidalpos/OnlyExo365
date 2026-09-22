# Architecture Overview

OnlyExo365 is a Windows-only desktop administration tool for Microsoft Exchange Online and Microsoft Graph.

## Runtime Model

The application operates across two dedicated processes communicating via Named Pipes:

- **`OnlyExo365.Shell` (`net10.0-windows`)**: WPF desktop presentation, configuration ingestion, local license catalog caching, and worker supervision.
- **`OnlyExo365.Worker` (`net10.0`)**: Out-of-process execution host powered by `Microsoft.PowerShell.SDK 7.6.1`. Manages Exchange and Graph runspaces, command execution, and error classification.
- **`OnlyExo365.Contracts` (`net10.0`)**: Shared binary contract defining IPC messages, DTOs, domain error taxonomy, and configuration contracts.

Details on pipe names, protocol framing, and session security are documented in [IPC Architecture](ipc.md).

## Internal Layering Rules

Responsibilities are partitioned internally across projects:

- **Domain**: Pure evaluators, capability rules, validation, and error models. Free of UI, storage, or PowerShell process dependencies.
- **Application**: Use-case orchestration and worker client interactions without direct WPF or PowerShell coupling.
- **Infrastructure**: IPC transport, DPAPI secret storage, persistent logs (`PersistentLogWriter`), configuration file loaders, and PowerShell runspaces.
- **Presentation**: WPF views, ViewModels, commands, and localized resources (`Loc`).

## Source Layout

- `src/OnlyExo365.Shell`: WPF presentation application.
- `src/OnlyExo365.Worker`: Out-of-process PowerShell execution engine.
- `src/OnlyExo365.Contracts`: Shared contracts and IPC abstractions.
- `tests/OnlyExo365.Tests`: Unit, characterization, and UI layout regression tests.
- `scripts/`: Canonical local entrypoints and automation workflows.
- `build/`: Architecture enforcement, coverage, security scans, and smoke test scripts.
- `installer/`: Inno Setup script for `OnlyExo365.Setup.exe`.
