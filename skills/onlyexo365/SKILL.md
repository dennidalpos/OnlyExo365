---
name: onlyexo365
description: Maintain or audit the OnlyExo365 Windows WPF shell and its out-of-process PowerShell worker. Use for repository changes involving IPC, Exchange or Graph operations, packaging, cleanup, or environment qualification.
---

# OnlyExo365

Work from the repository root. Before changing code, read `PROJECT_STATUS.json` for open work and [the development workflow](../../docs/development/workflow.md) for the current pins and commands. Treat historical audit counts as evidence for their recorded revisions, not current validation.

## Preserve these boundaries

- This is a Windows-only WPF shell with an out-of-process PowerShell worker. Keep approved exact-version module imports and the bootstrap Tools available if an import fails.
- Preserve the IPC contract: deadlines use `ErrorCode.Timeout`; voluntary cancellation uses `WasCancelled`; submitted tenant writes are not rolled back and must not be retried automatically.
- Pass the session token directly through the child environment. DPAPI protects distinct password references, not that token.
- Do not use the production tenant for offline validation. Dummy cmdlets and offline tests do not qualify Microsoft 365 behavior.
- Preserve pre-existing edits, ignored audit reports, inputs, and the fixtures recorded in `docs/audits/2026-10-10-retained-fixtures.md`.

## Choose validation that matches the change

- Run the smallest relevant native check before and after code changes. After edits, run `pwsh -NoProfile ./build/assert-architecture-constraints.ps1` and `git diff --check`; do not raise architecture hotspot limits.
- Use the documented `scripts/agents` compile and test commands. Restore/bootstrap only when their documented prerequisites require it.
- Treat installer smoke, tenant validation, and interactive UI/accessibility work as separate environment qualifications. Read [qualification guidance](../../docs/development/qualification.md) before attempting them.
- Installer qualification requires a disposable Windows host with no prior installation; signing or packaging success alone is insufficient. Preserve backup and manifest when recovery is incomplete.
- eDiscovery qualification requires delegated Interactive authentication on an interactive host; DeviceCode and app-only flows do not qualify it.

## Safe maintenance

- The clean script must preserve reports and inputs and reject tracked files or reparse points. Never stop operator processes to clear build locks.
- Do not commit or push unless the user explicitly requests it. Update tracker or documentation only when the work changes their truthful status.
- For source pins, package commands, and gate behavior, use [the workflow](../../docs/development/workflow.md) rather than duplicating them here.
