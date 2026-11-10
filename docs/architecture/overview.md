# Architecture

| Project | Responsibility |
| --- | --- |
| Shell (`net10.0-windows`) | WPF views/ViewModels, localization, configuration, local SKU cache, Worker supervision |
| Worker (`net10.0`) | Embedded PowerShell, Exchange/Graph runspaces, command mapping, capability detection and error classification |
| Contracts (`net10.0`) | IPC envelopes, DTOs, configuration, module policy, protected secret references and persistent logs |
| Tests (`net10.0-windows`) | Unit, generated-command, local-pipe, script and offscreen WPF regressions |

These responsibilities are folders and types inside the existing projects, not separate Domain/Application/Infrastructure assemblies. Shell `Bootstrap` composes pages and services; `UseCases` coordinates connection; Worker `Operations` dispatches requests into `PowerShell` modules. The 18 page ViewModels are constructed eagerly and retain application lifetime. The isolated bootstrap fixture took 152–157 ms for its first case and 4–21 ms for later cases in two local runs, including JIT, fixture setup and assertions. That proxy does not establish a full application startup bottleneck, so page lifetime/navigation semantics remain unchanged. Real startup profiling is part of UX-08; see the [local remediation log](../audits/2026-10-11-local-remediation.md).

The two processes communicate through [IPC](ipc.md). Catalog refresh runs inside Shell without IPC. The [module policy](../../src/OnlyExo365.Contracts/Data/PowerShellModuleBootstrapPolicy.json) is embedded in Contracts and shared by both hosts and tenant validation.

## Components and data paths

```mermaid
flowchart LR
    UI["Shell: WPF views and ViewModels"] --> Services["Shell: configuration, catalog and exports"]
    UI <-->|"Local named pipe"| Worker["Worker: operation dispatcher"]
    Worker --> Engine["Embedded PowerShell and runspaces"]
    Engine <-->|"Exchange module"| EXO["Exchange Online / Purview"]
    Engine <-->|"Graph modules"| Graph["Microsoft Graph"]
    Contracts["Contracts: DTOs, envelopes and policies"] -.-> UI
    Contracts -.-> Worker
    Services --> Disk["Local files"]
```

This describes the existing process and project boundaries. Shell's catalog download is a separate network path; Exchange/Graph administrative commands run in Worker.

## Request lifecycle

```mermaid
sequenceDiagram
    actor Operator
    participant Shell
    participant Pipe as Local named pipe
    participant Worker
    participant Cloud as Exchange or Graph
    Operator->>Shell: Select action and provide inputs
    Note over Operator,Shell: Mutation workflows ask for confirmation
    Shell->>Pipe: Request envelope with correlation ID
    Pipe->>Worker: Validated operation and payload
    Worker->>Cloud: Execute PowerShell command
    Worker-->>Pipe: Progress and log events when emitted
    Cloud-->>Worker: Result or service error
    Worker-->>Pipe: Response envelope
    Pipe-->>Shell: Typed result or normalized error
    Shell-->>Operator: Show result, warning or failure
    Note over Shell,Cloud: Timeout or cancellation does not roll back a cloud write
```

Cancellation and deadlines bound local request handling; they do not prove that a remote mutation was undone. The application resilience pipeline retries selected read operations only; it does not control retries internal to vendor modules. See [IPC](ipc.md) for framing, cancellation and message limits.

## Runtime decision

Retain .NET 10, an exact SDK pin with roll-forward disabled, and an embedded PowerShell SDK in the separate Worker to keep PowerShell execution outside WPF. Packaging targets `win-x64`; the current pinned versions are listed once in [workflow](../development/workflow.md). Dependency revisions and their evidence belong to the [critical audit](../audits/2026-10-10-critical-audit.md).

`scripts/` contains public entrypoints; `scripts/internal/common.ps1` provides shared helpers; `build/` implements gates, packaging and validation; `installer/OnlyExo365.iss` defines installation. Hotspot constraints are enforced by `build/assert-architecture-constraints.ps1`.
