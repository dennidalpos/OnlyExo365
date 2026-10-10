# Retained fixtures: CLEAN-01

Historical audit/remediation scratch remains because two bounded/literal cleanup attempts per stage were rejected by the execution harness: `rejected: blocked by policy`. The reported hypothesis was a direct recursive-delete restriction; no narrower reason was supplied. No subsequent deletion was attempted. This review preserves those fixtures and reports.

| Owned fixture | Retention reason |
| --- | --- |
| `artifacts/powershell-audit/scratch` | Initial probe copies/outputs; both guarded computed and verified literal removal rejected |
| `artifacts/powershell-remediation/checks`, `probes`, `icons`, `recovery-functions.ps1` | Remediation fixtures; two cleanup attempts rejected |
| `artifacts/powershell-regressions/a016e68d7b904ae2a2098485c9d44f39` | First tracked-output disposable Git clone; initial cleanup hit read-only Git object; future fixture cleanup clears read-only attributes |
| `%TEMP%/onlyexo365-localdata-backup-88b38d842a0e413ab929f1e5d2c5b5d1` | Fixture-only backup retained after initial cross-volume Directory.Move failure |
| `%TEMP%/onlyexo365-localdata-backup-c2a2bfed8e5546ecbb606e7e5c160842` | Same fixture-only failure; missing-entry fixture subsequently changed to rename within its backup root |
| `artifacts/documentation-review-20261010` | This review's comparison baseline, disposable Roslyn verifier/build outputs and IPC probe; literal removal rejected once by automatic approval review with blocked by policy; not retried |

Both temporary manifests were checked during remediation and pointed exclusively to owned scratch inputs, not personal/product data. The normal initial backup roundtrip's temporary backup was removed successfully. Production PowerShell move/restore regressions subsequently passed. The documentation-review directory was checked for containment, Git ignore, absence of tracked files/reparse points and 90 session-owned entries before its rejected removal.

Source repairs and checks are in the [PowerShell audit](2026-10-10-powershell-audit.md). CLEAN-01 remains pending in [PROJECT_STATUS.json](../../PROJECT_STATUS.json); verify ownership/manifest contents again before future cleanup, and do not retry under the unchanged rejected policy.
