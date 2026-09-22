# OnlyExo365 Agent Directives

Windows-only WPF application for Microsoft Exchange Online and Microsoft Graph administration with out-of-process PowerShell worker.

## Project Structure & Architecture

- **`OnlyExo365.Shell`** (`net10.0-windows`): WPF UI, configuration loader, SKU catalog updater, worker supervisor.
- **`OnlyExo365.Worker`** (`net10.0`): PowerShell host (`Microsoft.PowerShell.SDK 7.6.1`), IPC server, command dispatching.
- **`OnlyExo365.Contracts`** (`net10.0`): Shared IPC message envelopes, DTOs, domain error taxonomy.
- **`OnlyExo365.Tests`** (`net10.0-windows`): Unit and characterization tests.

## Quirks & Hard Rules

- **Strict SDK Pinning**: `global.json` pins .NET SDK `10.0.401` with `rollForward: disable`. Scripts check this strictly and will fail if the local machine does not have `10.0.401`.
- **Architectural File & Line Constraints**: `build/assert-architecture-constraints.ps1` enforces line limits on hotspot files (e.g. `ErrorClassifier.cs` <= 520, `AppCompositionRoot.cs` <= 80, `WorkerService.cs` <= 720). Any edit must not exceed these limits.
- **IPC Protocol**: Named Pipes (`OnlyExo365_IPC_Main`, `OnlyExo365_IPC_Events`) with UTF-8 line-delimited JSON (`\n`). Handshake uses DPAPI-protected token via `ONLYEXO365_IPC_SESSION_TOKEN`.
- **App Data Paths**: Persistent runtime data lives in `%LocalAppData%\OnlyExo365\` (`logs`, `ipc-secrets`, `exports`, `LicenseCatalog`).

## Verified Commands

```powershell
# Verify architecture line count constraints
pwsh ./build/assert-architecture-constraints.ps1
```
