# PowerShell remediation — 2026-10-10

Historical review of all 31 tracked production scripts, helpers, workflows/composite action, installer and verification paths. Initial inventory: **6,273 lines**; hashes in `artifacts/powershell-audit/inventory.json`. Generated PowerShell was exercised on existing regression paths, not every command (IMP-01). No production tenant, operator-data mutation or module/tool installation. Source findings PS-01..11 were subsequently repaired; environment qualification and retained fixture cleanup remain open in [PROJECT_STATUS.json](../../PROJECT_STATUS.json).

## Findings, reproduction and final behavior

Old line numbers/recommended patches were removed because they described superseded source. The original failed behavior and distinct evidence remain below.

| ID | Priority | Reproduced defect → verified repair |
| --- | --- | --- |
| PS-01 | P1 | Smoke finally deleted the only backup after injected restore failure; partial moves escaped the returned manifest → mutable recovery state, atomic entries before moves, hash-validated restore, separate smoke-created data, retained evidence on failure, empty-only backup removal. Existing installs rejected before backup. Native cases cover roundtrip/idempotence/corruption/partial move/missing entry/original-finally failure. |
| PS-02 | P1 | Dry-run accepted external recursive export deletion; broad bin/obj/log/cache/IDE cleanup ignored ownership → validate all targets, literal paths, generated roots/direct project outputs only; reject external/root/ancestor/reparse/tracked content. Preserve reports, coverage, IDE/input/unknown directories; explicit exports/imports restricted to matching artifacts roots. |
| PS-03 | P1 | String-prefix containment accepted artifacts-other; build could stop unrelated processes → removed arbitrary process termination/retry logic; bounded clean fails on locks. Smoke stops only its own process objects and propagates shutdown errors; fixture process exit/repeated/null cleanup pass. No operator process was stopped in reproduction. |
| PS-04 | P1 | Canonical policy preflight exit 1: `requiredModules` missing under strict mode → safe optional properties, malformed required entries rejected, failure report persisted. Optional Exchange/Graph argument dictionaries retain identity; minimal configuration/report cases pass. |
| PS-05 | P2 | Exclusively locked fixture remained while clean printed success/exit 0 → terminating deletions/native exit checks, completed-only statistics, nonzero failure. Locked/native-failure cases pass; -All additionally clears only NuGet HTTP cache. |
| PS-06 | P2 | Deterministic pack substitute removed security sentinel/run-1; in-memory hashes hid lost evidence → bounded clean retains prior evidence and both snapshots. Real two-build Release comparison: publish/package reproducible, empty mismatch list. |
| PS-07 | P2 | Tenant script had no Connect-IPPSSession, overclaiming Compliance → separate Compliance/search-only connections with cloud/delegated endpoints. DeviceCode/ManagedIdentity and app-only eDiscovery fail explicitly. Full offline AppCertificate fixture records both connections and six failed search probes; metadata is not live-write qualification. |
| PS-08 | P2 | Dummy certificate store selected a pre-existing same-subject certificate for deletion → remove exact run-owned thumbprint/private key, surface cleanup failure. Dummy-store preservation and real signing/timestamped-copy checks pass; source hashes/CurrentUser My inventory unchanged; no trust roots added. |
| PS-09 | P2 | Setup compiled before payload signing; final outer signature could hide unsigned embedded EXEs → sign payload, rebuild setup, sign setup, verify. Smoke compares installed hashes and, for signed setup, signer validity/equality. Mismatch regressions pass; trusted installed payload remains AUD-07. |
| PS-10 | P2 | One-SKU CSV produced OrderedDictionary / Entries:4 → preserve entries array, validate columns/GUID/product identity before atomic replacement. Quoted/one-entry/invalid cases pass; real Microsoft CSV generated 620 entries to separate evidence; tracked catalog unchanged. |
| PS-11 | P3 | Dead SDK-major/log helpers, unused smoke preference, six duplicate path resolvers, repeated scans/shared readers → removed confirmed dead/duplicate helpers and broad scans, retained active CLI wrappers. 6,273 → 5,821 lines; analyzer 191 → 146 warnings. Icon regeneration matches ICO/PNG SHA256; shared resolvers/readers exercised by native signing/security. Clean dry-run 3.49 s has no comparable baseline, so no speedup claim. |

Source entrypoints: [build](../../build), [shared helpers](../../scripts/internal/common.ps1), [catalog generator](../../scripts/agents/refresh-microsoft365-sku-catalog.ps1), [signed workflow](../../.github/workflows/release-signed.yml), [installer](../../installer/OnlyExo365.iss). Offline dummy cmdlets/certificates stay in [script-regressions.ps1](../../tests/powershell/script-regressions.ps1); the recovery/signing checks use owned disposable fixtures.

## Initial verification

Environment: Windows, PowerShell 7.6.6, SDK 10.0.401, PSScriptAnalyzer 1.25.0, Inno Setup 6, Windows SDK SignTool. Initial main checkout clean. First Debug no-build test attempt failed because only Release binaries existed; canonical Debug compilation then passed.

| Check | Recorded result before repair |
| --- | --- |
| AST parse / default analyzer | 31 scripts, 0 syntax errors / 191 warnings, 0 errors |
| Debug compile / full suite | 0 errors / 789 passed, 0 failed/skipped, 15 s |
| Doctor, Inno/SignTool discovery, architecture/diff | Passed |
| Actual tenant preflight, SkipGraph/no module installation | Exit 1, optional requiredModules error; no connection |
| Isolated behavior/source probes | 11 satisfied, 9 violated; unresolved violations made harness nonzero |

Satisfied probes: nested repo/RID resolution, SHA256 backup roundtrip, reproducibility/report, runtime major/patch, quoted uninstall path, both coverage-threshold sides, scratch icon, quoted CSV, four dummy authentication modes, release assets twice with archive/count/hash checks. Nine violations map to PS-01/02/03/04/05/06/07/08/10; PS-09 was source ordering. Backup failure injection executed original finally; reproducibility used a substitute with canonical cleanup semantics; certificate discovery used dummy store commands.

Initial analyzer warnings: Write-Host 129, singular nouns 21, ShouldProcess 12, default-true switch 8, unused parameter 6, approved verb 6, empty catch 5, automatic-variable assignment 3, plaintext password 1. WorkerExecutablePath/RequiredParameters are read in nested scriptblocks, not confirmed dead. Avoid CLI/logging renames merely to silence diagnostics; signing password interface remains a separate contract concern. No suppressions/bypass flags were added.

Initial evidence under `artifacts/powershell-audit/`: `inventory.json`, `analyzer.json`, `probes.json`, `probes.log`, `tenant-preflight.log`, `tests/onlyexo365-tests.trx`, `offline-probes.ps1`.

## Final recorded verification

| Check | Result after repair |
| --- | --- |
| Debug compile / full suite | 0 errors / 810 passed, 0 failed/skipped |
| Release test-project build / full suite | 0 warnings/errors / 811 passed, 0 failed/skipped |
| Final focused script/security cases | 54 passed after helper/caller repairs |
| Actual publish-path resolution | Passed |
| Native two-build Release packaging | publish_reproducible=true, package_reproducible=true, mismatches=[]; both snapshots/prior evidence retained |
| Disposable native signing/timestamp/verification | Passed; original hashes and certificate inventory unchanged |
| PowerShell AST / production analyzer | 32 scripts including harness, 0 parse errors / 146 warnings, 0 errors, no suppressions |
| Architecture / NuGet / gitleaks | Passed / no known vulnerabilities across 4 projects / no leaks |

Harness PSScriptAnalyzer failed twice: `Reference Position should begin before start Position of Range.` Static analyzer coverage of that harness is **not verified**; parser and 21 native-script regressions pass. Stop retrying that unchanged tool failure.

Final evidence under `artifacts/powershell-remediation/`: `tests/onlyexo365-tests.trx`, `tests-release/onlyexo365-tests.trx`, `remediation-final-targeted.trx`, `inventory.json`, `analyzer-production-final.json`, `signing-validation.json`/`.log`, `tenant-offline-coverage.json`, `catalog-current.json`, `clean-dry-run.log`. Native reports/snapshots: `artifacts/reproducibility/` and `artifacts/security/`. Commands are consolidated in [workflow](../development/workflow.md), [packaging](../releases/packaging.md) and [signing](../releases/signing.md).

Full gate, operator-host real installer and live tenant writes were not run. Unsigned reproducibility excludes timestamped signed builds. Release was regenerated at this stage, before the subsequent [MainWindow startup fix](2026-10-10-ui-ux-audit.md); regenerate again before AUD-07. Real installer/tenant/UI checks are in [qualification](../development/qualification.md); exact blocked scratch paths are in [retained fixtures](2026-10-10-retained-fixtures.md).

References: [strict mode](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/set-strictmode?view=powershell-7.6), [error/native exit handling](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_error_handling?view=powershell-7.6), [literal paths](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.management/remove-item?view=powershell-7.6), [Compliance/eDiscovery](https://learn.microsoft.com/en-us/powershell/exchange/connect-to-scc-powershell?view=exchange-ps), [collection return](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_return?view=powershell-7.6), [analyzer rules](https://learn.microsoft.com/en-us/powershell/utility-modules/psscriptanalyzer/rules-recommendations?view=ps-modules), [Inno SignTool](https://jrsoftware.org/ishelp/topic_setup_signtool.htm), [Microsoft SignTool](https://learn.microsoft.com/en-us/windows/win32/seccrypto/signtool).
