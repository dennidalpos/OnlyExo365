# Development, Build, and Test Workflow

This document owns local development commands and CI gate mapping.

## Baseline

- Windows x64
- PowerShell 7+ (`pwsh`)
- .NET SDK `10.0.203` exactly; `global.json` disables roll-forward
- NuGet lockfiles enabled
- runtime identifier `win-x64`
- Inno Setup 6 only for packaging

## Canonical Commands

Run from the repository root:

```powershell
pwsh ./scripts/bootstrap.ps1 -RuntimeIdentifier win-x64
pwsh ./scripts/build.ps1 -Configuration Debug -RuntimeIdentifier win-x64
pwsh ./scripts/start.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBuild
pwsh ./scripts/clean.ps1
pwsh ./scripts/gate.ps1 -RuntimeIdentifier win-x64
```

Packaging prerequisites:

```powershell
pwsh ./scripts/agents/doctor.ps1 -CheckPackaging
pwsh ./scripts/Install-InnoSetup.ps1
pwsh ./scripts/Install-InnoSetup.ps1 -Install -PackageManager Auto
```

`scripts/Install-InnoSetup.ps1` checks `INNOSETUP_BIN`, `INNOSETUP_HOME`, `C:\Program Files (x86)\Inno Setup 6`, and `C:\Program Files\Inno Setup 6`. With `-Install`, it tries winget first and Chocolatey second unless `-PackageManager` selects one explicitly.

## Scripts Layout

- `scripts/bootstrap.ps1`: restores repository dependencies.
- `scripts/build.ps1`: runs the canonical build without publishing.
- `scripts/start.ps1`: starts the WPF shell from source.
- `scripts/clean.ps1`: removes generated repository outputs.
- `scripts/gate.ps1`: cleans repository outputs and user-local app state, checks prerequisites, compiles, tests, scans, and builds distributable artifacts. Use `-InstallPrerequisites` to install Inno Setup when missing, `-CleanPerMachineAppSettings` to also remove shared ProgramData app settings, and `-RunReproducibility` / `-RunSigningValidation` for release-adjacent checks.
- `scripts/pack.ps1`: builds publish output and creates `OnlyExo365.Setup.exe`.
- `scripts/Install-InnoSetup.ps1`: checks or installs the packaging prerequisite.
- `scripts/agents/*.ps1`: CI, release, test, and maintenance automation entrypoints.
- `scripts/internal/common.ps1`: shared PowerShell helpers for repository scripts.

## Fresh-Install Development Flow

1. Install PowerShell 7+.
2. Install .NET SDK `10.0.203`; other SDK versions fail because `global.json` disables roll-forward.
3. From the repository root, run `pwsh ./scripts/bootstrap.ps1 -RuntimeIdentifier win-x64`.
4. Run `pwsh ./scripts/build.ps1 -Configuration Debug -RuntimeIdentifier win-x64`.
5. Run `pwsh ./scripts/agents/test.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBootstrap`.
6. For packaging only, install or configure Inno Setup 6, then run `pwsh ./scripts/pack.ps1 -Configuration Release -LockedMode -RuntimeIdentifier win-x64`.

Framework-dependent publish output is the default and requires the .NET 10 Desktop Runtime on target machines. Self-contained publish output is available through `-SelfContained` on `scripts/pack.ps1`, `scripts/gate.ps1`, or `build/build.ps1`.

## Local CI-Equivalent Validation

```powershell
pwsh ./scripts/agents/doctor.ps1 -CheckPackaging
pwsh ./scripts/bootstrap.ps1 -LockedMode:$true -RuntimeIdentifiers win-x64
pwsh ./scripts/agents/compile.ps1 -Configuration Debug -RuntimeIdentifiers win-x64 -NoBootstrap
pwsh ./scripts/agents/test.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBootstrap
pwsh ./build/assert-architecture-constraints.ps1
pwsh ./build/assert-no-vulnerable-packages.ps1 -SolutionPath OnlyExo365.sln -ReportPath artifacts/security/nuget-vulnerabilities.json
pwsh ./build/run-secret-scan.ps1 -SourcePath . -ReportPath artifacts/security/gitleaks.sarif
dotnet test tests/OnlyExo365.Tests/OnlyExo365.Tests.csproj -c Debug --no-build --no-restore --artifacts-path artifacts/build --collect:"XPlat Code Coverage" --results-directory artifacts/test-results --filter "FullyQualifiedName~OnlyExo365.Tests.WorkerCommandTests|FullyQualifiedName~OnlyExo365.Tests.MailboxReportingCommandTests|FullyQualifiedName~OnlyExo365.Tests.IpcContractsTests|FullyQualifiedName~OnlyExo365.Tests.IpcSecretHandlingTests|FullyQualifiedName~OnlyExo365.Tests.PowerShellModuleBootstrapPolicyTests|FullyQualifiedName~OnlyExo365.Tests.PersistentLogWriterTests|FullyQualifiedName~OnlyExo365.Tests.OperationDispatcherMessagingTests"
$coveragePath = Get-ChildItem artifacts/test-results -Recurse -Filter coverage.cobertura.xml | Sort-Object LastWriteTimeUtc | Select-Object -Last 1 -ExpandProperty FullName
pwsh ./build/assert-code-coverage.ps1 -CoveragePath $coveragePath -MinimumLineCoveragePercent 20 -IncludePackage OnlyExo365.Worker,OnlyExo365.Contracts
pwsh ./scripts/pack.ps1 -Configuration Release -LockedMode -RuntimeIdentifiers win-x64
```

There is no separate repository-level lint or standalone typecheck command beyond the compile/build gates.

## CI Mapping

`.github/actions/windows-release-baseline/action.yml` is the shared baseline used by the release workflows. It performs:

- SDK setup from `global.json`
- Inno Setup installation
- `doctor`
- locked restore
- Debug compile
- unit tests
- coverage enforcement
- architecture checks
- NuGet vulnerability scan
- secret scan
- release packaging
- package reproducibility
- signing-flow validation

Release assets come from `scripts/pack.ps1` plus `scripts/agents/publish.ps1`; there is no separate manual release evidence script.
