# Development workflow

## Pinned baseline

| Component | Repository requirement / source |
| --- | --- |
| Platform | Windows x64; packaging defaults to win-x64 |
| .NET SDK | 10.0.401 exactly; [global.json](../../global.json), roll-forward disabled |
| Target frameworks | .NET 10; Windows target for Shell/Tests |
| Worker PowerShell SDK | 7.6.6; [Worker project](../../src/OnlyExo365.Worker/OnlyExo365.Worker.csproj) |
| External PowerShell | Scripts declare 7.0; Exchange bootstrap and tenant validation require 7.6+ |
| Approved modules | EXO 3.10.1; Graph Authentication, Users, Users.Actions, Identity.DirectoryManagement 2.41.1; [canonical policy](../../src/OnlyExo365.Contracts/Data/PowerShellModuleBootstrapPolicy.json) |
| Installer compiler | Inno Setup 6; shared release CI pins 6.7.1 |

## Local commands

Run from repository root. Restore precedes compile/test when outputs are absent:

```powershell
pwsh ./scripts/agents/doctor.ps1 -CheckPackaging
pwsh ./scripts/bootstrap.ps1 -RuntimeIdentifier win-x64
pwsh ./scripts/build.ps1 -Configuration Debug -RuntimeIdentifier win-x64
pwsh ./scripts/start.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBuild
pwsh ./scripts/agents/compile.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBootstrap
pwsh ./scripts/agents/test.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBootstrap -NoBuild -NoRestore
pwsh ./build/assert-architecture-constraints.ps1
pwsh ./build/assert-no-vulnerable-packages.ps1 -SolutionPath OnlyExo365.sln -ReportPath artifacts/security/nuget-vulnerabilities.json
pwsh ./build/run-secret-scan.ps1 -SourcePath . -ReportPath artifacts/security/gitleaks.sarif
```

The test runner accepts `-ResultsDirectory`; select focused cases directly through `dotnet test --filter` with `--artifacts-path artifacts/build`, as release CI does. The architecture gate checks required files and hotspot line limits. No dedicated formatter is configured; use `git diff --check` for whitespace. Coverage enforcement is in the shared release CI action, with a 20% Worker/Contracts minimum for its selected backend tests.

## Clean and gate

```powershell
pwsh ./scripts/clean.ps1 -DryRun
pwsh ./scripts/clean.ps1
pwsh ./scripts/gate.ps1 -RuntimeIdentifier win-x64
```

Clean validates all targets before deletion: `artifacts/build`, `artifacts/publish`, `artifacts/packages`, and direct .NET project bin/obj under src/tests. It rejects tracked files, reparse points, external/root/ancestor targets. Reports, coverage, IDE files, unknown directories and default exports/imports survive. Explicit export/import cleanup must stay under the matching artifacts directory. `-All` additionally clears NuGet HTTP cache. Locks and native failures return nonzero; build does not stop operator processes.

Gate performs clean `-All`, prerequisites, locked restore, Debug compilation/tests, architecture/security and Release packaging. Reproducibility and disposable signing validation require `-RunReproducibility` / `-RunSigningValidation`. Gate does not invoke installer smoke, tenant validation or interactive UI qualification. It checks packaging prerequisites even when packaging is skipped; inspect [gate.ps1](../../scripts/gate.ps1) before choosing flags.

After an approved dependency change, regenerate lockfiles through NuGet:

```powershell
dotnet restore OnlyExo365.sln --force-evaluate --artifacts-path artifacts/build --runtime win-x64
```

`scripts/Install-InnoSetup.ps1` discovers the compiler; `-Install` requests installation. Packaging/assets, catalog maintenance and environment qualification are documented in [packaging](../releases/packaging.md), [catalog](../licensing/catalog.md) and [qualification](qualification.md).
