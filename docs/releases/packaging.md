# Release and Packaging

Procedures for packaging, smoke testing, digital signing, and publishing OnlyExo365 releases.

## Packaging

Build the distributable installer:

```powershell
pwsh ./scripts/pack.ps1 -Configuration Release -LockedMode -RuntimeIdentifier win-x64
```

Outputs produced:
- `artifacts/publish/win-x64/`: Published application binaries.
- `artifacts/packages/OnlyExo365.Setup.exe`: Inno Setup installer executable.

> [!NOTE]
> By default, packaging produces framework-dependent binaries requiring the .NET 10 Desktop Runtime on target machines. Pass `-SelfContained` to include the runtime within the package.

---

## Smoke Testing

Execute local end-to-end smoke verification:

```powershell
pwsh ./build/run-smoke-tests.ps1
```

Validates:
- Binary completeness in publish directories.
- Successful headless launch of `OnlyExo365.Shell.exe`.
- Out-of-process worker spawn and IPC handshake establishment.
- Creation of UI, Worker, and Supervisor log files.
- Worker console show/hide toggle behavior.

---

## Code Signing

Digital signing scripts reside in `build/`:
- `sign-artifacts.ps1`: Applies Authenticode signatures to binaries and installers.
- `verify-signatures.ps1`: Validates file signatures against certificate roots.
- `validate-artifact-signing.ps1`: Gate asserting all published artifacts are signed.

---

## Release Asset Generation

To bundle release artifacts for distribution:

```powershell
pwsh ./scripts/agents/publish.ps1 -ReleaseTag <vX.Y.Z> -RuntimeIdentifier win-x64
```

Creates under `artifacts/publish/release-assets/`:
- `OnlyExo365-<tag>-win-x64-publish.zip`
- `OnlyExo365-<tag>-win-x64-setup.exe`
- `OnlyExo365-<tag>-win-x64.sha256`
- `OnlyExo365-<tag>-win-x64-assets.json`
