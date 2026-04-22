# PowerShell SDK and .NET Baseline

## Context

The worker depends on `Microsoft.PowerShell.SDK` `7.6.0`, while the product ships as a Windows desktop shell plus a separate PowerShell-backed worker. The repository also enforces packaging, smoke validation, signing-flow validation, and release automation around that split runtime model.

## Decision

The repository baseline is:

- `.NET SDK` `10.0.203` pinned in `global.json`
- `OnlyExo365.Contracts`: `net10.0`
- `OnlyExo365.Worker`: `net10.0`
- `OnlyExo365.Shell`: `net10.0-windows`
- `OnlyExo365.Tests`: `net10.0-windows`
- `Microsoft.PowerShell.SDK` `7.6.0` in the worker
- Windows/x64-only packaging and release outputs

The previous split between `Application`, `Infrastructure`, and `Domain` projects is intentionally retired. The solution now keeps only the runtime projects and the shared contracts needed by the final product architecture.

## Rationale

- the worker-facing toolchain must close end-to-end, not only at compile time
- a single framework baseline is simpler to verify and maintain
- the final architecture is clearer when the shell and worker own their real responsibilities directly
- the repository no longer carries compatibility layers that existed only for historical structure
