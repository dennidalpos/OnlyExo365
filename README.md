![OnlyExo365 app icon](src/OnlyExo365.Shell/Assets/AppIcon.png)

# OnlyExo365

OnlyExo365 is a Windows-only WPF desktop application for Exchange Online administration. The repository is intentionally consolidated around one runtime architecture:

- `OnlyExo365.Shell`: WPF operator shell
- `OnlyExo365.Worker`: separate PowerShell-backed worker process
- `OnlyExo365.Contracts`: shared IPC contracts, DTOs, configuration contracts, and diagnostics
- `OnlyExo365.Tests`: automated regression and repository-perimeter tests

## Local Baseline

- Windows
- PowerShell 7
- .NET SDK `10.0.203` from `global.json`
- Inno Setup 6 only when packaging `OnlyExo365.Setup.exe`

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

`scripts/gate.ps1` is the local repository gate. It cleans generated repository outputs and user-local OnlyExo365 app data, checks prerequisites, optionally installs packaging prerequisites, runs locked bootstrap, compiles, tests, runs validation scans, publishes, and creates the setup EXE package. Use `-CleanPerMachineAppSettings` only when shared ProgramData app settings should also be removed.

## Technical Documentation

- [Architecture overview](docs/architecture/overview.md)
- [Development, build, and test workflow](docs/development/build-test.md)
- [Operator setup and runtime configuration](docs/operations/operator-setup.md)
- [Release and verification](docs/releases/verification.md)
- [Microsoft 365 SKU catalog maintenance](docs/maintenance/licensing-catalog.md)
- [PowerShell and .NET baseline decision](docs/decisions/2026-04-11-powershell-sdk-major-upgrade-strategy.md)

## License

OnlyExo365 is distributed under the proprietary terms in [LICENSE](LICENSE).
