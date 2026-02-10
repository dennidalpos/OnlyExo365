# ExchangeAdmin

ExchangeAdmin is a .NET 8 solution for administering Microsoft Exchange environments. The solution includes:

- **ExchangeAdmin.Presentation**: a WPF desktop UI targeting `net8.0-windows` that operators use to manage mailboxes, distribution lists, and system status.  
- **ExchangeAdmin.Worker**: a console worker targeting `net8.0` that executes background or delegated tasks (for example, PowerShell-based operations).  
- **ExchangeAdmin.Application / Domain / Infrastructure / Contracts**: shared layers containing application logic, domain models, infrastructure services, and contracts used by both the UI and the worker.  

## Prerequisites

- **Windows** (required for WPF and the `net8.0-windows` UI target).
- **.NET SDK 8.x**.
- **PowerShell 7.x** (the worker references `Microsoft.PowerShell.SDK`).

## Solution structure

```
ExchangeAdmin.sln
src/
  ExchangeAdmin.Application/
  ExchangeAdmin.Contracts/
  ExchangeAdmin.Domain/
  ExchangeAdmin.Infrastructure/
  ExchangeAdmin.Presentation/   # WPF UI (net8.0-windows)
  ExchangeAdmin.Worker/         # Console worker (net8.0)
```

## Build

From the repository root:

```bash
dotnet build ExchangeAdmin.sln
```

### Notes

- The WPF UI project includes a build target that copies the worker output (`ExchangeAdmin.Worker`) into the UI output folder after a successful build. This ensures the UI can locate the worker at runtime.  

## Run (UI)

```bash
dotnet run --project src/ExchangeAdmin.Presentation/ExchangeAdmin.Presentation.csproj
```

## Run (Worker)

```bash
dotnet run --project src/ExchangeAdmin.Worker/ExchangeAdmin.Worker.csproj
```


## Build & clean scripts

Repository scripts live in `build/`:

- `build/clean.ps1`: removes artifacts, `bin/obj`, temporary files, test/coverage outputs, and optionally cache-related data.
- `build/build.ps1`: validates prerequisites, restores/builds the solution, publishes binaries, and can optionally produce an MSI.

Quick examples:

```powershell
# Dry-run cleanup
pwsh ./build/clean.ps1 -DryRun

# Full cleanup including cache-related items
pwsh ./build/clean.ps1 -All

# Build release + publish self-contained for a specific RID
pwsh ./build/build.ps1 -Configuration Release -Publish -SelfContained -RuntimeIdentifier win-x64
```

## UI checkbox/filter behavior analysis

A detailed analysis of each user-facing checkbox/filter (trigger behavior, backend mapping, and side-effects) is available in:

- `docs/checkbox-filter-logic.md`

## Development tips

- **UI + Worker pairing**: if you run the UI project, ensure the worker project builds successfully so the post-build copy target can place the worker binaries alongside the UI output.
- **Project references**: application, domain, infrastructure, and contracts are shared; changes in these layers affect both the UI and worker.
- **Mailbox permissions and quotas**: SendAs entries are displayed with friendly names while removal actions use the original trustee identity for compatibility with `Remove-RecipientPermission`; mailbox quota bytes are parsed from localized PowerShell size strings in the worker.

## Troubleshooting

| Symptom | Likely cause | Fix |
| --- | --- | --- |
| WPF project fails to build | Non-Windows OS or missing Windows SDK | Build on Windows with .NET 8 SDK installed |
| Worker fails at runtime | Missing PowerShell runtime | Install PowerShell 7.x and ensure it is available on PATH |
