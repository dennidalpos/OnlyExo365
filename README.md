# ExchangeAdmin

## Overview
ExchangeAdmin is a Windows WPF application that manages Exchange Online operations through an out-of-process PowerShell worker. The UI communicates with the worker over named pipes using a shared IPC contract so long-running tasks can stream progress and logs while keeping the UI responsive.

## Requirements
- Windows (WPF target).
- .NET 8 SDK for building (or .NET 8 runtime if using non-self-contained publish).
- PowerShell 7+ for the worker process.
- ExchangeOnlineManagement PowerShell module installed on the machine running the worker.
- WiX Toolset v3.14 only if you want to build the MSI installer.

## Installation
1. Ensure the requirements above are installed.
2. Build the solution:
   ```powershell
   dotnet build ExchangeAdmin.sln
   ```

## Configuration
The application reads the following environment variables:

| Variable | Purpose | Values |
| --- | --- | --- |
| `EXCHANGEADMIN_EXO_ENV` | Sets the Exchange Online environment used by the worker when connecting. | `O365Default`, `O365GermanyCloud`, `O365USGovGCCHigh`, `O365USGovDoD`, `O365China` |
| `EXCHANGEADMIN_DISABLE_EXO` | Disables Exchange Online connections in the UI when set to a truthy value. | `1`, `true`, `yes`, `on` (case-insensitive). |

## Running
### Development
Build and run the WPF application (it will start the worker process when needed):
```powershell
dotnet run --project src/ExchangeAdmin.Presentation/ExchangeAdmin.Presentation.csproj
```

### Production / Publish
Use the build script to compile and publish (self-contained by default):
```powershell
pwsh build/build.ps1 -Configuration Release -Publish
```

For a quick local build/publish/launch flow:
```powershell
pwsh build/quick-test.ps1
```

## Tests
There are no automated tests included in the repository. Use `dotnet build` or the build scripts above to validate local builds.

## Lint / Format
No linting or formatting scripts are defined in the repository.

## Project Structure
- `src/ExchangeAdmin.Presentation` — WPF UI project (entry point for the desktop app).
- `src/ExchangeAdmin.Application` — Application services and use cases bridging UI and worker IPC.
- `src/ExchangeAdmin.Infrastructure` — IPC client/supervisor that manages the worker process and messaging.
- `src/ExchangeAdmin.Worker` — PowerShell worker that executes Exchange Online commands.
- `src/ExchangeAdmin.Contracts` — Shared IPC messages and DTOs between UI and worker.
- `src/ExchangeAdmin.Domain` — Shared domain errors and results used across layers.
- `build/` — Build and quick-test PowerShell scripts.
- `installer/` — WiX installer definition for MSI packaging.

## Key Commands
```powershell
# Build the full solution
 dotnet build ExchangeAdmin.sln

# Run the WPF app
 dotnet run --project src/ExchangeAdmin.Presentation/ExchangeAdmin.Presentation.csproj

# Publish (self-contained default)
 pwsh build/build.ps1 -Configuration Release -Publish

# Quick publish & launch
 pwsh build/quick-test.ps1
```

## Troubleshooting
- **Worker cannot connect or module missing**: Install the ExchangeOnlineManagement module in PowerShell 7 and restart the app.
- **Exchange connections disabled**: Clear `EXCHANGEADMIN_DISABLE_EXO` or set it to `0`, then restart the app.
- **Execution policy issues**: The worker attempts to set the current user execution policy to `RemoteSigned` on startup if required; check PowerShell policy if startup warnings persist.
