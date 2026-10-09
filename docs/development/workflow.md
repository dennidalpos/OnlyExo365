# Development Workflow

Canonical development, compilation, testing, and quality-gate procedures.

## Environment Baseline

- **OS**: Windows x64
- **Shell**: PowerShell 7+ (`pwsh`)
- **.NET SDK**: `10.0.401` exactly, pinned in `global.json` with `rollForward: disable`
- **Packaging**: Inno Setup 6 (required only for installer generation)
- **Target Runtime**: `win-x64` (`net10.0` and `net10.0-windows`)

---

## Canonical Commands

Run from the repository root:

```powershell
# 1. Restore dependencies with locked NuGet mode
pwsh ./scripts/bootstrap.ps1 -RuntimeIdentifier win-x64

# 2. Compile solution
pwsh ./scripts/build.ps1 -Configuration Debug -RuntimeIdentifier win-x64

# 3. Launch application from source
pwsh ./scripts/start.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBuild

# 4. Clean artifacts and outputs
pwsh ./scripts/clean.ps1

# 5. Run full local gate (clean, build, test, architecture, security, pack)
pwsh ./scripts/gate.ps1 -RuntimeIdentifier win-x64

# 6. Build release installer package
pwsh ./scripts/pack.ps1 -Configuration Release -LockedMode -RuntimeIdentifier win-x64
```

> [!NOTE]
> `scripts/gate.ps1` cleans repository build outputs only. Personal logs, configuration, secrets, catalogs, and exports are preserved automatically. The former `KeepLocalAppData` and `CleanPerMachineAppSettings` flags were removed.

---

## Scripts Layout

- `scripts/bootstrap.ps1`: Restores solution dependencies with NuGet lockfile enforcement.
- `scripts/build.ps1`: Builds all projects without publishing.
- `scripts/start.ps1`: Runs the WPF shell directly from build outputs.
- `scripts/clean.ps1`: Cleans repository build artifacts (`artifacts/build`, `artifacts/publish`).
- `scripts/gate.ps1`: Local end-to-end CI pipeline equivalent.
- `scripts/pack.ps1`: Publishes and builds the Inno Setup distribution package.
- `scripts/Install-InnoSetup.ps1`: Automated discovery or installation of Inno Setup 6.
- `scripts/agents/*.ps1`: Automation entrypoints:
  - `doctor.ps1`: Environment and toolchain verification.
  - `compile.ps1`: Compilation helper.
  - `test.ps1`: Test execution runner.
  - `publish.ps1`: Release asset packaging and SHA256 checksum generation.
  - `refresh-microsoft365-sku-catalog.ps1`: M365 licensing CSV downloader and catalog generator.

---

## Quality Gates & Verification

The repository enforces architectural and security boundaries:

- **Architecture Constraints**:
  ```powershell
  pwsh ./build/assert-architecture-constraints.ps1
  ```
  Validates maximum line count thresholds on core orchestrators and checks presence of mandatory architecture files.
- **Vulnerability Scan**:
  ```powershell
  pwsh ./build/assert-no-vulnerable-packages.ps1 -SolutionPath OnlyExo365.sln -ReportPath artifacts/security/nuget-vulnerabilities.json
  ```
- **Secret Scan**:
  ```powershell
  pwsh ./build/run-secret-scan.ps1 -SourcePath . -ReportPath artifacts/security/gitleaks.sarif
  ```
- **Test Suite**:
  ```powershell
  pwsh ./scripts/agents/test.ps1 -Configuration Debug -RuntimeIdentifier win-x64
  ```

## Offline Audit Verification

After locked restore, compile and run the offline suite without invoking the full gate:

```powershell
pwsh ./scripts/agents/compile.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBootstrap
pwsh ./scripts/agents/test.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBootstrap -NoBuild -NoRestore
```

Use external PowerShell 7.6+ for `build/run-tenant-validation.ps1` and interactive Exchange bootstrap. The module policy is embedded from `src/OnlyExo365.Contracts/Data/PowerShellModuleBootstrapPolicy.json`; Shell, Worker, and validation scripts select exact approved versions. Regenerate NuGet lockfiles with `dotnet restore OnlyExo365.sln --force-evaluate --artifacts-path artifacts/build --runtime win-x64` when approved dependencies change.
