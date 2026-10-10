# OnlyExo365 Agent Directives

Windows-only WPF Shell with an out-of-process PowerShell Worker. Never use the production tenant for offline audit validation.

## Operational constraints

- Read `PROJECT_STATUS.json` for open work; audit counts describe their historical revisions. Source pins and commands are consolidated in `docs/development/workflow.md`.
- Preserve pre-existing edits and ignored audit reports/inputs. Never commit/push without current authorization.
- Keep approved exact-version module imports and bootstrap Tools available after import failure.
- Keep IPC envelopes: deadlines use `ErrorCode.Timeout`; voluntary cancellation uses `WasCancelled`. Submitted writes cannot be rolled back by cancellation; never retry writes automatically.
- Session token passes directly through the child environment. DPAPI protects separate password references, not that token.
- Clean preserves reports/inputs and rejects tracked files/reparse points. Never stop operator processes to clear build locks.
- Installer smoke requires a disposable Windows host without an existing installation; packaging/signature-copy success does not qualify uninstall or trusted installed payloads. Incomplete recovery retains backup/manifest.
- eDiscovery validation requires delegated Interactive authentication and an interactive host. Offline dummy cmdlets do not qualify Microsoft 365 behavior.
- `docs/audits/2026-10-10-retained-fixtures.md` lists earlier cleanup rejected twice by policy. Preserve those paths; do not retry under unchanged policy.
- Run architecture constraints after edits; do not raise hotspot limits.

## Verified in this documentation review (2026-10-10)

```powershell
pwsh -NoProfile ./scripts/agents/compile.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBootstrap
pwsh -NoProfile ./build/assert-architecture-constraints.ps1
git diff --check
```

Environment qualification: `docs/development/qualification.md`. Source/evidence reconciliation: `docs/audits/2026-10-10-documentation-review.md`.
