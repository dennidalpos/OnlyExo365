# Environment qualification

These checks are separate from offline tests and the local gate. Installer and tenant qualification were removed from the active task list at user request on 2026-10-11 because the required environments are unavailable. Their quality gates remain unverified in [PROJECT_STATUS.json](../../PROJECT_STATUS.json); the procedures are retained for future qualification. Previous results are in the three [audit reports](../../README.md#documentation).

## Disposable installer host

Regenerate Release payload/setup from current source before qualification. Use a disposable Windows VM without an existing installation:

```powershell
pwsh ./scripts/pack.ps1 -Configuration Release -LockedMode -RuntimeIdentifier win-x64
pwsh ./build/run-smoke-tests.ps1 -PublishPath artifacts/publish -SetupExePath artifacts/packages/OnlyExo365.Setup.exe -ReportPath artifacts/smoke/smoke-report.json
```

Smoke checks publish completeness, offline Shell/Worker startup, authenticated IPC, logs and console toggle, then real install/reinstall/uninstall by default. It rejects an existing installation before backing up default product data. `-SkipRealInstall` omits installer qualification. Reinstall checks installed Shell/Worker hashes against publish and, for signed setup, valid signatures with the same signer. Check preserved exports/unknown files and removed installed payload.

Backup entries/hashes are written atomically before moves. Restore validates backups, separates smoke-created data and verifies original hashes. Any incomplete recovery fails and retains backup/manifest; only an empty recovered root is removed. Packaging or disposable signing alone does not qualify this flow. See [signing](../releases/signing.md).

## Dedicated test tenant

External PowerShell 7.6+, exact approved modules and a dedicated test tenant are required; never use production for audit validation. Configure [settings](../configuration/settings.md) and inspect [run-tenant-validation.ps1](../../build/run-tenant-validation.ps1) before execution.

The script samples Exchange, audit/search, mail-security and Graph reads; it inspects write-command metadata without performing search/purge/hold writes. It writes a failure report even when preflight fails. Compliance and search-only sessions are separate; unsupported authentication yields failed probes and nonzero exit. Interactive fallback to DeviceCode cannot qualify eDiscovery. Cloud/delegated endpoints follow configuration.

The script does not exercise every application feature, Message Trace or cancellation. Complete real approved imports/authentication, progressive Trace and failure-propagation qualification separately. Offline dummy cmdlets demonstrate control flow only; authentication constraints are in [authentication](../operations/authentication.md).

## UX-08: interactive Windows UI

Use 900×600 and 1200×800 windows, 100/150/200% DPI and IT/EN. Check keyboard navigation, Escape, owner modality, Narrator, real high contrast and realistic long/populated data. Profile realized rows and responsiveness on Dashboard/Compliance/Security. Offscreen WPF layout/virtualization probes do not qualify assistive technology or production performance.
