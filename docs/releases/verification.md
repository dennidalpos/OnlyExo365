# Release and Verification

This document owns packaging, smoke validation, signing checks, release asset preparation, and tenant validation. Development commands and CI gate mapping are covered in [../development/build-test.md](../development/build-test.md).

## Release Outputs

`pwsh ./scripts/pack.ps1 -Configuration Release -LockedMode` creates:

- `artifacts/publish/win-x64/`: framework-dependent Windows application payload for 64-bit Windows.
- `artifacts/publish/win-x86/`: framework-dependent Windows application payload for 32-bit Windows.
- `artifacts/packages/OnlyExo365.Setup.exe`: canonical Inno Setup installer.

`pwsh ./scripts/publish.ps1 -ReleaseTag <tag>` materializes durable release assets under `artifacts/publish/release-assets`:

- `OnlyExo365-<tag>-multi-arch-publish.zip`
- `OnlyExo365-<tag>-multi-arch-setup.exe`
- `OnlyExo365-<tag>-multi-arch.sha256`
- `OnlyExo365-<tag>-multi-arch-assets.json`

Build, packaging, smoke validation, signing, release upload, and tenant validation are separate steps.

## Packaging Prerequisites

Local setup EXE packaging requires:

- Windows
- PowerShell 7
- .NET SDK `10.0.202`, as pinned in `global.json`
- versioned NuGet lockfiles
- Inno Setup 6

Check the packaging toolchain:

```powershell
pwsh ./scripts/doctor.ps1 -CheckPackaging
pwsh ./scripts/Install-InnoSetup.ps1
```

Install Inno Setup locally when needed:

```powershell
pwsh ./scripts/Install-InnoSetup.ps1 -Install
```

The Inno Setup discovery path used by `doctor.ps1`, `pack.ps1`, and `Install-InnoSetup.ps1` is `PATH`, `INNOSETUP_BIN`, `INNOSETUP_HOME`, then default Inno Setup 6 install directories.

## Packaging Behavior

Installer authoring lives in `installer/ExchangeAdmin.iss`.

Current setup behavior:

- installation scope: per-machine
- privileges: administrator
- architecture: one setup EXE with x64 payload selection on 64-bit Windows and x86 payload selection on 32-bit Windows
- default install directory: `C:\Program Files\OnlyExo365`
- installed payload: files from `artifacts/publish/win-x64/` on 64-bit Windows or `artifacts/publish/win-x86/` on 32-bit Windows
- shortcuts: common Start Menu and common Desktop
- registry: install location plus current-user default log, secret, and export locations
- uninstall cleanup: installed app directory, default logs, DPAPI IPC secrets, and default exports
- upgrade note: the installer AppId remains unchanged and `UsePreviousAppDir=yes` preserves an existing ExchangeAdmin-branded install root during upgrade

## Smoke Validation

`build/run-smoke-tests.ps1` verifies:

- required publish files exist
- .NET runtime prerequisites match the generated runtimeconfig files
- the publish payload can launch `ExchangeAdmin.Presentation.exe`
- the worker process starts from the expected path
- UI, supervisor, and worker logs are written
- the setup EXE can install to a temporary root when real install validation is enabled
- the installed app can launch
- uninstall registry, payload cleanup, residual cleanup, and service cleanup complete
- pre-existing local app data is backed up and restored during installer validation

Real installer validation requires a Windows host without a conflicting existing `ExchangeAdmin` installation. The smoke harness now also reinstalls the setup over the existing install root and fails if more than one uninstall registry entry is registered for the product.

## Signing

Signing scripts and helpers live under `build/`:

- `sign-artifacts.ps1`
- `verify-signatures.ps1`
- `validate-artifact-signing.ps1`
- `signing-helpers.ps1`
- `resolve-signtool.ps1`

Local unsigned packaging is not a signed release. A signed release requires the signing workflow or equivalent local signing material plus successful Authenticode signature verification.

The signed release workflow requires:

- `WINDOWS_SIGNING_PFX_BASE64`
- `WINDOWS_SIGNING_PFX_PASSWORD`
- optional `WINDOWS_SIGNING_CERT_SHA1` for signer thumbprint verification

## CI Release Paths

- `.github/workflows/release-validation.yml`: shared Windows release baseline, smoke validation, and artifact upload.
- `.github/workflows/release-signed.yml`: shared baseline, signing, signature verification, smoke validation, release asset materialization, and GitHub release upload.
- `.github/workflows/release-tenant-validation.yml`: manual read-only validation against a real tenant and certificate-backed app identity.

Tenant validation requires these GitHub secrets:

- `EXCHANGEADMIN_APP_ID`
- `EXCHANGEADMIN_EXO_ORGANIZATION`
- `EXCHANGEADMIN_CERT_THUMBPRINT`
- `EXCHANGEADMIN_CERTIFICATE_PFX_BASE64`
- `EXCHANGEADMIN_CERTIFICATE_PFX_PASSWORD`
- `EXCHANGEADMIN_GRAPH_TENANT_ID`, unless Graph validation is skipped for the manual run

## Operational Limits

- Build success is not packaging success.
- Packaging success is not release readiness.
- Unsigned local artifacts are not signed release artifacts.
- Smoke validation can alter install state and local operator data during the test, then restore backed-up data.
- Tenant validation is not locally reproducible without a real tenant and required certificate material.
- GitHub release upload is performed only by the signed release workflow path.
