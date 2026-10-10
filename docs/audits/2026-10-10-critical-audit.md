# Critical remediation — 2026-10-10

Historical offline reliability/data-protection evidence. Local runspaces/pipes and dummy cmdlets reproduced defects without production tenant access or module installation on the operator profile. Later UI/script results supersede these verification counts; current open work is in [PROJECT_STATUS.json](../../PROJECT_STATUS.json).

## Findings and repairs

| ID | Priority | Repaired behavior / evidence |
| --- | --- | --- |
| AUD-01 | P1 | Gate personal-state cleanup flags removed; uninstall preserves exports/unknown files. Source regressions pass; isolated uninstall remains AUD-07. |
| AUD-02 | P1 | Named credential parameters bind correctly. Failed mailbox/deleted-mailbox/contact/resource/device/public-folder/migration/group-member listings throw before mapping; deprecation text does not swallow errors. Dummy cmdlets verify binding/failure paths. |
| AUD-03 | P1 | Request lifetime starts before queue admission; saturation promptly returns ServiceUnavailable while heartbeat/cancel remain responsive. Local-pipe cases prove queued cancellation/deadline prevents dispatch; timeout differs from cancellation. No automatic write retry. |
| AUD-04 | P2 | Server Retry-After survives local cap; cancellation returns a structured result even during exception backoff. A 60-second delay with a 1-ms cap stays pending until cancelled. |
| AUD-05 | P2 | Trace accumulates requested page plus one lookahead with partial/exact counts. Empty/short/multibatch/stalled-cursor and invalid date/page cases pass. |
| AUD-06 | P2 | One exact-version module policy for both hosts, licensing and tenant validation. Newer installed modules do not substitute; missing approved versions fail clearly; Tools survives EXO import failure. |

Initial 14 data-protection/PowerShell/retry regressions failed before repair; pipe/Trace probes also reproduced defects. Exception-backoff probes exposed cancellation escaping structured results. Baseline success did not cover those defects.

## Dependency revision

Retained .NET 10 / SDK 10.0.401. Worker SDK 7.6.1 → 7.6.6; EXO 3.9.2 → 3.10.1; four Graph modules 2.35.1 → 2.41.1; ProtectedData/Xml → 10.0.12. OpenXml/test dependencies unchanged. NuGet regenerated lockfiles; locked restore passed. These are repository pins, not a claim of latest releases.

Original sources: [PowerShell release](https://github.com/PowerShell/PowerShell/releases/tag/v7.6.6), [EXO Gallery notes](https://www.powershellgallery.com/packages/ExchangeOnlineManagement/3.10.1), [Graph releases](https://github.com/microsoftgraph/msgraph-sdk-powershell/releases), [ProtectedData](https://www.nuget.org/packages/System.Security.Cryptography.ProtectedData/10.0.12), [Xml](https://www.nuget.org/packages/System.Security.Cryptography.Xml/10.0.12). Behavioral references: [Graph throttling](https://learn.microsoft.com/en-us/graph/throttling), [Trace V2](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-messagetracev2), [Exchange prerequisites](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2), [Inno uninstall guidance](https://jrsoftware.org/ishelp/topic_uninstalldeletesection.htm).

## Recorded verification

| Check | Result at this revision |
| --- | --- |
| Baseline / final full Debug suite | 733 / 768 passed; final 0 failed/skipped, 35 new regressions |
| Targeted cases | 186 passed, then 3 retry checks after exception-backoff repair |
| Debug compile / architecture | Passed, 0 errors; IPC server 684/760 non-empty lines |
| Changed PowerShell syntax | 4 scripts, 0 errors |
| NuGet / gitleaks | No known vulnerabilities across 4 projects / no leaks |
| Release compile, publish and installer creation | Passed; framework-dependent win-x64 |
| git diff --check | Passed |

Evidence: `artifacts/test-results/onlyexo365-tests.trx`, `artifacts/security/nuget-vulnerabilities.json`, `artifacts/packages/OnlyExo365.Setup.exe`. Commands are consolidated in [workflow](../development/workflow.md) and [packaging](../releases/packaging.md). Full gate, signing and reproducibility were not run at this stage; later [PowerShell remediation](2026-10-10-powershell-audit.md) records the subsequent signing/reproducibility results.

AUD-07/AUD-08 remain environment qualification, detailed in [qualification](../development/qualification.md). Timeout/cancellation does not roll back submitted writes. IMP-01 generated-script coverage, IMP-02 atomic exports and IMP-03 standard CSV parsing/wider localization remain proposals; the [UI audit](2026-10-10-ui-ux-audit.md) covers the completed client-localization portion.
