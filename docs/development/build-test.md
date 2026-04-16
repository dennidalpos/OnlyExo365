# Development, Build, and Test Workflow

This document owns local development commands, test commands, and CI gate mapping. Runtime operator configuration is covered in [../operations/operator-setup.md](../operations/operator-setup.md). Packaging and release readiness are covered in [../releases/verification.md](../releases/verification.md).

## Baseline

The repository is Windows-first.

Required local baseline:

- Windows
- PowerShell 7
- .NET SDK `10.0.202`, pinned by `global.json`
- versioned NuGet lockfiles

Inno Setup 6 is only required when creating the setup EXE locally.

## Canonical Commands

Run commands from the repository root.

```powershell
pwsh ./scripts/bootstrap.ps1
pwsh ./scripts/doctor.ps1
pwsh ./scripts/compile.ps1 -Configuration Debug
pwsh ./scripts/test.ps1 -Configuration Debug -RuntimeIdentifier win-x64
pwsh ./scripts/clean.ps1
```

Packaging prerequisite checks:

```powershell
pwsh ./scripts/doctor.ps1 -CheckPackaging
pwsh ./scripts/Install-InnoSetup.ps1
```

Local Inno Setup installation, when needed:

```powershell
pwsh ./scripts/Install-InnoSetup.ps1 -Install
```

There is no repository-level lint, format, or standalone typecheck command beyond the compile/build gates.

## Build Artifacts

Repository scripts use `artifacts/` for build, test, security, packaging, smoke, and release outputs. That directory is ignored by `.gitignore`.

`scripts/compile.ps1` builds the solution with `/warnaserror` and the configured artifacts path.

`scripts/test.ps1` runs the solution tests and writes TRX output under `artifacts/test-results` by default.

## CI Gate Mapping

The shared CI baseline lives in `.github/actions/windows-release-baseline/action.yml`.

It performs:

- .NET SDK setup from `global.json`
- Inno Setup 6 installation
- SignTool prerequisite resolution
- `scripts/doctor.ps1 -CheckPackaging`
- locked restore through `scripts/bootstrap.ps1`
- Debug compile through `scripts/compile.ps1`
- NuGet vulnerability scan through `build/assert-no-vulnerable-packages.ps1`
- repository secret scan through `build/run-secret-scan.ps1`
- automated tests through `scripts/test.ps1`
- filtered coverage collection and enforcement through `build/assert-code-coverage.ps1`
- architecture constraints through `build/assert-architecture-constraints.ps1`
- Release packaging through `scripts/pack.ps1`
- package reproducibility through `build/verify-package-reproducibility.ps1`
- artifact signing-flow validation through `build/validate-artifact-signing.ps1`

`release-validation.yml` adds smoke validation and artifact upload after the shared baseline. `release-signed.yml` adds signing, signature verification, release asset materialization, and GitHub release upload. `release-tenant-validation.yml` runs a manual tenant-backed validation path.

## Local CI-Equivalent Checks

The closest local equivalent to the shared CI baseline is:

```powershell
pwsh ./scripts/Install-InnoSetup.ps1
pwsh ./scripts/doctor.ps1 -CheckPackaging
pwsh ./scripts/bootstrap.ps1
pwsh ./scripts/compile.ps1 -Configuration Debug -NoBootstrap
pwsh ./build/assert-no-vulnerable-packages.ps1 -SolutionPath ExchangeAdmin.sln -ReportPath artifacts/security/nuget-vulnerabilities.json
pwsh ./build/run-secret-scan.ps1 -SourcePath . -ReportPath artifacts/security/gitleaks.sarif
pwsh ./scripts/test.ps1 -Configuration Debug -ResultsDirectory artifacts/test-results/unit -NoBootstrap -NoBuild -NoRestore
pwsh ./build/assert-architecture-constraints.ps1
pwsh ./scripts/pack.ps1 -Configuration Release -Clean:$false -LockedMode
pwsh ./build/verify-package-reproducibility.ps1 -Configuration Release
pwsh ./build/validate-artifact-signing.ps1 -Path artifacts/publish,artifacts/packages
pwsh ./build/run-smoke-tests.ps1 -PublishPath artifacts/publish -SetupExePath artifacts/packages/OnlyExo365.Setup.exe -ReportPath artifacts/smoke/smoke-report.json
```

`scripts/bootstrap.ps1`, `scripts/compile.ps1`, `scripts/pack.ps1`, and `build/build.ps1` now restore or package both `win-x64` and `win-x86` by default so the repository produces one unified Windows setup EXE with architecture-specific payloads under `artifacts/publish/win-x64` and `artifacts/publish/win-x86`.

Coverage collection in CI uses `dotnet test --collect:"XPlat Code Coverage"` with a filtered set of core backend tests before `build/assert-code-coverage.ps1` enforces the current coverage floor.
