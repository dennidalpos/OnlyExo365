# OnlyExo365 Agent Directives

Windows-only WPF Shell with an out-of-process PowerShell Worker. Never use the production tenant for offline audit validation.

## Repository Quirks

- `global.json` requires .NET SDK `10.0.401` exactly, with roll-forward disabled.
- Worker embeds PowerShell `7.6.6`; external interactive Exchange bootstrap and tenant validation require PowerShell `7.6+`.
- The canonical approved module policy is embedded from `src/OnlyExo365.Contracts/Data/PowerShellModuleBootstrapPolicy.json`. Shell and Worker must import exact approved versions; bootstrap Tools remain available after import failure.
- `build/assert-architecture-constraints.ps1` enforces hotspot line limits. Run it after edits; do not raise limits to accommodate changes.
- IPC is authenticated, UTF-8 line-delimited JSON over named pipes. `ONLYEXO365_IPC_SESSION_TOKEN` contains the DPAPI-protected session token. Keep existing envelopes; deadlines use `ErrorCode.Timeout`, voluntary cancellation uses `WasCancelled`.
- Persistent state is under `%LocalAppData%\OnlyExo365`. The gate preserves personal state automatically. Uninstall preserves exports and unknown installation files, removing only installed payload and transient logs/IPC secrets.
- Installer smoke tests require an isolated disposable host. Packaging success does not verify install/uninstall.

## Commands Verified on 2026-10-10

```powershell
pwsh ./scripts/agents/compile.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBootstrap
pwsh ./scripts/agents/test.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBootstrap -NoBuild -NoRestore
pwsh ./build/assert-architecture-constraints.ps1
pwsh ./build/assert-no-vulnerable-packages.ps1 -SolutionPath OnlyExo365.sln -ReportPath artifacts/security/nuget-vulnerabilities.json
pwsh ./scripts/pack.ps1 -Configuration Release -Clean:$false -LockedMode -RuntimeIdentifier win-x64
```

Current evidence and pending checks: `docs/audits/2026-10-10-critical-audit.md`, `docs/audits/2026-10-10-ui-ux-audit.md` and `PROJECT_STATUS.json`.
