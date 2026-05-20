![OnlyExo365 app icon](src/OnlyExo365.Shell/Assets/Generated/AppIcon.png)

# OnlyExo365

OnlyExo365 is a Windows-only WPF desktop application for Exchange Online administration. The repository is intentionally consolidated around one runtime architecture:

- `OnlyExo365.Shell`: WPF operator shell
- `OnlyExo365.Worker`: separate PowerShell-backed worker process
- `OnlyExo365.Contracts`: shared IPC contracts, DTOs, configuration contracts, and diagnostics
- `OnlyExo365.Tests`: automated regression and repository-perimeter tests

## Verified Baseline

- Windows x64
- PowerShell 7+ (`pwsh`)
- .NET SDK `10.0.204` exactly, pinned by `global.json` with roll-forward disabled
- NuGet lockfiles in each project
- Runtime identifier `win-x64`
- Inno Setup 6 only when creating `OnlyExo365.Setup.exe`

## Canonical Commands

Run from the repository root:

```powershell
pwsh ./scripts/bootstrap.ps1 -RuntimeIdentifier win-x64
pwsh ./scripts/build.ps1 -Configuration Debug -RuntimeIdentifier win-x64
pwsh ./scripts/start.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBuild
pwsh ./scripts/clean.ps1
pwsh ./scripts/gate.ps1 -RuntimeIdentifier win-x64
pwsh ./scripts/pack.ps1 -Configuration Release -LockedMode -RuntimeIdentifier win-x64
```

`scripts/gate.ps1` is the local repository gate. It cleans generated repository outputs and, unless `-KeepLocalAppData` is passed, user-local OnlyExo365 app data. It then checks prerequisites, runs locked bootstrap, compiles, tests, runs repository validation scans, publishes, and creates the setup EXE package. Use `-InstallPrerequisites` to install Inno Setup through winget or Chocolatey when packaging prerequisites are missing. Use `-CleanPerMachineAppSettings` only when shared ProgramData app settings should also be removed.

## Fresh Install Notes

For development on a clean Windows machine, install PowerShell 7 and the pinned .NET SDK before running `scripts/bootstrap.ps1`. Packaging also needs Inno Setup 6, discovered from `INNOSETUP_BIN`, `INNOSETUP_HOME`, or the standard install directories. Runtime packages are framework-dependent by default and require .NET 10 Desktop Runtime; pass `-SelfContained` to `scripts/pack.ps1` or `scripts/gate.ps1` to produce self-contained publish output.

At application runtime, Exchange and Graph work require PowerShell 7 plus Microsoft 365 access. The app can check and bootstrap the required PowerShell modules from the Tools page: `ExchangeOnlineManagement 3.9.2` and Microsoft Graph modules based on `Microsoft.Graph.Authentication 2.35.1`.

## Technical Documentation

- [Architecture overview](docs/architecture/overview.md)
- [Development, build, and test workflow](docs/development/build-test.md)
- [Operator setup and runtime configuration](docs/operations/operator-setup.md)
- [Release and verification](docs/releases/verification.md)
- [Microsoft 365 SKU catalog maintenance](docs/maintenance/licensing-catalog.md)
- [PowerShell and .NET baseline decision](docs/decisions/2026-04-11-powershell-sdk-major-upgrade-strategy.md)

## License

OnlyExo365 is distributed under the proprietary terms in [LICENSE](LICENSE).
