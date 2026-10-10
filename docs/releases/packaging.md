# Packaging and release assets

Requirements are in [workflow](../development/workflow.md). Run from repository root:

```powershell
pwsh ./scripts/pack.ps1 -Configuration Release -LockedMode -RuntimeIdentifier win-x64
```

Outputs: `artifacts/publish/win-x64/` and `artifacts/packages/OnlyExo365.Setup.exe`. Default packaging is framework-dependent and requires .NET 10 Desktop Runtime; `-SelfContained` includes the runtime. Pack defaults to `Clean=false`, but publishing rebuilds the publish tree; prepare release assets afterward.

For unsigned reproducibility:

```powershell
pwsh ./build/verify-package-reproducibility.ps1 -Configuration Release -RuntimeIdentifier win-x64
```

The two snapshots/reports are retained under `artifacts/reproducibility`; prior test/security evidence survives bounded clean. Timestamped signed-package reproducibility is outside this comparison.

Prepare an already-packaged release bundle:

```powershell
pwsh ./scripts/agents/publish.ps1 -ReleaseTag vX.Y.Z -RuntimeIdentifier win-x64
```

Replace `vX.Y.Z` with the intended tag. Under `artifacts/publish/release-assets/`: `OnlyExo365-<tag>-win-x64-publish.zip`, `-setup.exe`, `.sha256`, `-assets.json`. The helper prepares files; release workflows perform GitHub upload.

## Installation data contract

[OnlyExo365.iss](../../installer/OnlyExo365.iss) removes installed payload, default logs and IPC secrets while preserving exports and unknown installation files. It does not indiscriminately delete the application directory, consistent with [Inno Setup guidance](https://jrsoftware.org/ishelp/topic_uninstalldeletesection.htm).

Compilation/reproducibility does not verify install/uninstall. Use [isolated qualification](../development/qualification.md) for recovery and installed-payload checks; [signing](signing.md) for the signed release order. AUD-07 remains pending and current Release artifacts must include the MainWindow startup repair.
