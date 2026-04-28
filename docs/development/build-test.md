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
pwsh ./scripts/bootstrap.ps1 -RuntimeIdentifier win-x64
pwsh ./scripts/build.ps1 -Configuration Debug -RuntimeIdentifier win-x64
pwsh ./scripts/start.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBuild
pwsh ./scripts/clean.ps1
```

Packaging prerequisites:

```powershell
pwsh ./scripts/agents/doctor.ps1 -CheckPackaging
pwsh ./scripts/Install-InnoSetup.ps1
```

## Scripts Layout

- `scripts/bootstrap.ps1`: restores repository dependencies.
- `scripts/build.ps1`: runs the canonical build without publishing.
- `scripts/start.ps1`: starts the WPF shell from source.
- `scripts/clean.ps1`: removes generated repository outputs.
- `scripts/pack.ps1`: builds publish output and creates `OnlyExo365.Setup.exe`.
- `scripts/Install-InnoSetup.ps1`: checks or installs the packaging prerequisite.
- `scripts/agents/*.ps1`: CI, release, test, and maintenance automation entrypoints.
- `scripts/internal/common.ps1`: shared PowerShell helpers for repository scripts.
- `build/release-baseline.ps1`: manual clean-worktree release evidence path; it is not the CI publication path.

## Local CI-Equivalent Validation

```powershell
pwsh ./scripts/agents/doctor.ps1 -CheckPackaging
pwsh ./scripts/bootstrap.ps1 -LockedMode:$true -RuntimeIdentifiers win-x64
pwsh ./scripts/agents/compile.ps1 -Configuration Debug -RuntimeIdentifiers win-x64 -NoBootstrap
pwsh ./scripts/agents/test.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBootstrap
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

`build/release-baseline.ps1` is intentionally separate from this CI baseline. Use it only when a human needs a clean-worktree evidence bundle under `artifacts/release` and, optionally, an annotated baseline tag. CI release assets still come from `scripts/pack.ps1` plus `scripts/agents/publish.ps1`.
