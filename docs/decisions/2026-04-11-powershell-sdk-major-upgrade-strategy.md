# Microsoft.PowerShell.SDK Major Upgrade and .NET 10 Baseline

## Context

The repository showed that a direct package-only upgrade from `Microsoft.PowerShell.SDK` `7.4.x` to `7.6.x` was not compatible with the earlier baseline:

- the `net8.0` worker stopped resolving `System.Management.Automation` after the lockfile update
- the `7.6.0` metapackage exposes compile assets starting at `net10.0`
- `ExchangeAdmin.Presentation` and `ExchangeAdmin.Tests` both reference `ExchangeAdmin.Worker`, so the change could not be isolated to the worker project alone

The repository also uses a Windows packaging baseline with an Inno Setup EXE, a framework-dependent publish output, and real installer smoke validation. The upgrade therefore had to close end-to-end, not just at worker compile time.

## Decision

The repository adopts the following verified combination:

- `global.json` pinned to `.NET SDK` `10.0.202`
- `src/ExchangeAdmin.Worker` retargeted to `net10.0`
- `src/ExchangeAdmin.Presentation` and `tests/ExchangeAdmin.Tests` retargeted to `net10.0-windows`
- shared libraries (`Application`, `Contracts`, `Domain`, `Infrastructure`) kept on `net8.0`
- `Microsoft.PowerShell.SDK` upgraded to `7.6.0` in the worker with aligned versioned lockfiles

The explicit choice was to retarget only the projects that must reference the worker directly, rather than retargeting the entire solution without evidence that it was necessary.

## Alternatives considered

### Upgrade only the package on `net8.0`

Rejected. This was the incompatible combination that broke worker compilation.

### Retarget the entire solution to `.NET 10`

Rejected as the first move. The shared `net8.0` libraries are consumable by the `net10.0` projects in this repository and did not need retargeting to close the actual task.

### Replace the metapackage with ad hoc granular PowerShell references

Rejected. It would increase drift, packaging risk, and maintenance cost without any supporting repository convention.

### Stay on `7.4.13`

Rejected as the end state. It kept the repository stable but left a real compatibility task open.

## Rationale

The repository uses a WPF desktop app, a separate worker, versioned lockfiles, an Inno Setup EXE, and real setup smoke validation. The major PowerShell SDK upgrade therefore required alignment across:

- worker compilation and the projects that reference it
- locked restore for those projects
- automated tests
- packaging and setup smoke validation

The most conservative repository-consistent change was a focused `.NET 10` retarget of only the worker-facing projects, leaving the rest of the solution unchanged where no evidence required further retargeting.

## Impact

- the worker now runs `Microsoft.PowerShell.SDK` `7.6.0` on `net10.0`
- the framework-dependent payload requires .NET Desktop Runtime 10 on operator machines
- shared libraries remain on `net8.0` until a future change is justified by real repository evidence
- operational tracking should not keep a separate PowerShell SDK upgrade task for this completed baseline

## Costs and limits

- development and packaged execution now depend on the .NET 10 SDK/runtime baseline where applicable
- any future retarget of the shared libraries should be justified by actual repository need, not cosmetic uniformity
