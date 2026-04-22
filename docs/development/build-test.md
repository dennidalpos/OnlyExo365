# Development, Build, and Test Workflow

This document owns local development commands and CI gate mapping.

## Baseline

- Windows
- PowerShell 7
- .NET SDK `10.0.203`
- NuGet lockfiles enabled
- Inno Setup 6 only for packaging

## Canonical Commands

Run from the repository root:

```powershell
pwsh ./scripts/doctor.ps1
pwsh ./scripts/bootstrap.ps1 -RuntimeIdentifier win-x64
pwsh ./scripts/compile.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBootstrap
pwsh ./scripts/test.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBootstrap
pwsh ./scripts/clean.ps1
```

Packaging prerequisites:

```powershell
pwsh ./scripts/doctor.ps1 -CheckPackaging
pwsh ./scripts/Install-InnoSetup.ps1
```

## Local CI-Equivalent Validation

```powershell
pwsh ./scripts/doctor.ps1 -CheckPackaging
pwsh ./scripts/bootstrap.ps1 -LockedMode:$true -RuntimeIdentifiers win-x64
pwsh ./scripts/compile.ps1 -Configuration Debug -RuntimeIdentifiers win-x64 -NoBootstrap
pwsh ./scripts/test.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBootstrap
pwsh ./build/assert-architecture-constraints.ps1
pwsh ./build/assert-no-vulnerable-packages.ps1 -SolutionPath OnlyExo365.sln -ReportPath artifacts/security/nuget-vulnerabilities.json
pwsh ./build/run-secret-scan.ps1 -SourcePath . -ReportPath artifacts/security/gitleaks.sarif
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
