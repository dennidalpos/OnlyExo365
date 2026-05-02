# Release and Verification

This document owns packaging, smoke validation, signing checks, and release asset preparation.

## Release Outputs

`pwsh ./scripts/pack.ps1 -Configuration Release -LockedMode -RuntimeIdentifier win-x64` produces framework-dependent release outputs:

- `artifacts/publish/win-x64/`
- `artifacts/packages/OnlyExo365.Setup.exe`

Pass `-SelfContained` to `scripts/pack.ps1` only when the release needs to carry the .NET runtime instead of requiring .NET 10 Desktop Runtime on the target machine.

`pwsh ./scripts/agents/publish.ps1 -ReleaseTag <tag> -RuntimeIdentifier win-x64` produces release assets under `artifacts/publish/release-assets`:

- `OnlyExo365-<tag>-win-x64-publish.zip`
- `OnlyExo365-<tag>-win-x64-setup.exe`
- `OnlyExo365-<tag>-win-x64.sha256`
- `OnlyExo365-<tag>-win-x64-assets.json`

## Packaging Behavior

Installer authoring lives in `installer/OnlyExo365.iss`.

Current setup behavior:

- architecture: x64 only
- installation scope: per-machine
- privileges: administrator
- default install directory: `C:\Program Files\OnlyExo365`
- shortcuts: common Start Menu and common Desktop
- license page: repository `LICENSE`
- uninstall cleanup: install directory plus default logs, IPC secrets, and exports

## Smoke Validation

`build/run-smoke-tests.ps1` validates:

- publish payload completeness
- .NET runtime prerequisites against runtimeconfig files
- launch of `OnlyExo365.Shell.exe`
- worker spawn from the expected path
- UI, supervisor, and worker log creation
- worker console toggle behavior from `Tools`
- worker console close command protection
- setup EXE install/reinstall/uninstall on a temporary root when real install validation is enabled

## Signing

Signing helpers live under `build/`:

- `sign-artifacts.ps1`
- `verify-signatures.ps1`
- `validate-artifact-signing.ps1`
- `signing-helpers.ps1`
- `resolve-signtool.ps1`

Unsigned local packaging is not a signed release.

## CI Release Paths

- `.github/workflows/release-validation.yml`
- `.github/workflows/release-signed.yml`
- `.github/workflows/release-tenant-validation.yml`

Tenant validation requires real certificate material and tenant credentials; it is not locally reproducible without them.

## Release Asset Preparation

Durable GitHub release assets are produced by `scripts/pack.ps1` and `scripts/agents/publish.ps1`.
