# Critical Audit and Remediation â€” 2026-10-10

## Scope and Baseline

Implemented the approved reliability/data-protection plan using local runspaces and named pipes. No production tenant operations or module installation on the operator profile were performed. The offline implementation was completed without commit/push; the user subsequently authorized saving the completed work on `main` and pushing it for session handoff.

Before changes: 733 tests passed, Debug build and architecture constraints passed, and NuGet reported no known vulnerabilities. Positive baseline results did not cover the behavioral defects below. Regression tests were added before fixes: the initial 14 data-protection/PowerShell/retry tests failed; IPC and progressive trace probes also reproduced failures. New exception-backoff probes exposed cancellation escaping the structured result.

## Corrected Findings

| ID | Priority | Change and evidence |
|---|---|---|
| AUD-01 | P1 | Gate personal-state cleanup and its two flags removed. Installer preserves exports and unknown application-directory files. Source checks pass; isolated uninstall remains pending. |
| AUD-02 | P1 | Credential scripts declare named parameters; dummy local cmdlets verify actual binding. Errors mentioning deprecation remain failures. Failed mailbox, deleted mailbox, contact, resource, device, public-folder, migration and group-member listings throw before mapping. |
| AUD-03 | P1 | Request lifetime starts before queue admission. Saturation returns `ServiceUnavailable` promptly while heartbeat/cancel remain responsive. Local-pipe tests verify queued cancellation/deadline prevents execution and timeout differs from voluntary cancellation. Writes are not retried automatically. |
| AUD-04 | P2 | Server `Retry-After` survives the local cap. A 60-second delay with a 1-ms local cap remains pending until cancellation; exception-backoff cancellation returns a structured cancelled result. |
| AUD-05 | P2 | Message Trace loads the requested page plus one record, stops immediately, and exposes partial/exact counts. Empty, short, multi-batch and stalled-cursor cases pass. Invalid page/date ranges fail before execution. |
| AUD-06 | P2 | Shared policy selects exact approved module versions for discovery/import in both hosts, licensing and tenant validation. Local cmdlets prove a newer installed version is not selected and a missing approved version fails clearly. Tools remain available after Exchange import failure. |

## Dependency Decisions and Current Primary Sources

- Worker `Microsoft.PowerShell.SDK`: 7.6.1 â†’ **7.6.6**, retaining .NET 10 and pinned SDK 10.0.401. [Official release](https://github.com/PowerShell/PowerShell/releases/tag/v7.6.6).
- `ExchangeOnlineManagement`: 3.9.2 â†’ **3.10.1**. Certificate-authentication fixes and the search-only Compliance fix justify the update; external PowerShell hosts require 7.6+. [Microsoft Gallery release notes](https://www.powershellgallery.com/packages/ExchangeOnlineManagement/3.10.1).
- All four approved Graph modules: 2.35.1 â†’ **2.41.1**. Includes authentication dependency isolation and re-import fixes. [Official releases](https://github.com/microsoftgraph/msgraph-sdk-powershell/releases).
- `System.Security.Cryptography.ProtectedData` and `System.Security.Cryptography.Xml`: **10.0.12** for baseline maintenance, without attributing a vulnerability to the previous versions. [ProtectedData](https://www.nuget.org/packages/System.Security.Cryptography.ProtectedData/10.0.12), [Xml](https://www.nuget.org/packages/System.Security.Cryptography.Xml/10.0.12).
- OpenXml and test framework versions are unchanged. NuGet regenerated lockfiles; locked restore passes.

Behavioral references: [Graph throttling](https://learn.microsoft.com/en-us/graph/throttling) requires respecting server retry delays; [Get-MessageTraceV2](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-messagetracev2) supports 90-day retention, at most 10 days per query, and up to 5000 records per batch; [Inno Setup](https://jrsoftware.org/ishelp/topic_uninstalldeletesection.htm) warns against deleting the whole application directory indiscriminately. [Exchange prerequisites](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2) inform supported OS requirements. Sources checked on 2026-10-10.

## Offline Verification

| Verification | Result |
|---|---|
| Targeted affected tests | 186/186 passed, followed by 3/3 retry checks after the exception-backoff fix |
| One complete suite | **768/768 passed, 0 failed, 0 skipped**; 35 additional regressions |
| Debug native compilation | Passed, 0 errors, project analyzers retained |
| Architecture | Passed; IPC server 684/760 non-empty lines |
| PowerShell syntax | 4 changed scripts parsed without errors |
| NuGet scan including transitive dependencies | No known vulnerabilities across 4 projects |
| Secret scan | gitleaks: no leaks found |
| Release compilation/publish/installer creation | Passed; framework-dependent `win-x64` package |
| Diff hygiene | `git diff --check` passed |

Commands:

```powershell
dotnet restore OnlyExo365.sln --force-evaluate --artifacts-path artifacts/build --runtime win-x64
pwsh ./scripts/agents/compile.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBootstrap
pwsh ./scripts/agents/test.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBootstrap -NoBuild -NoRestore
pwsh ./build/assert-architecture-constraints.ps1
pwsh ./build/assert-no-vulnerable-packages.ps1 -SolutionPath OnlyExo365.sln -ReportPath artifacts/security/nuget-vulnerabilities.json
pwsh ./scripts/pack.ps1 -Configuration Release -Clean:$false -LockedMode -RuntimeIdentifier win-x64
pwsh ./build/run-secret-scan.ps1 -SourcePath . -ReportPath artifacts/security/gitleaks.sarif
```

Evidence: `artifacts/test-results/onlyexo365-tests.trx`, `artifacts/security/nuget-vulnerabilities.json`, `artifacts/packages/OnlyExo365.Setup.exe`. The full gate was not run: its separate live/interactive checks were not part of offline validation.

## Pending Qualification

- **AUD-07 / P1:** Install/uninstall the new package in a disposable Windows VM. Check removal of installed files and survival of exports/unknown files with unchanged contents. The smoke script now asserts this contract. No isolated host was available in this session; this behavior is **not verified interactively**.
- **AUD-08 / P2:** Validate approved real modules, authentication modes, paging and error propagation against a dedicated test tenant. Local simulated cmdlets prove control flow only. Microsoft 365 validation remains separate; no production tenant was used.
- Signing and reproducibility checks were not requested or executed. Timeout/cancellation cannot roll back Microsoft 365 changes already submitted.

## Separate Improvement Proposals

1. **IMP-01 / P2:** Extend actual generated-script execution coverage beyond the regression paths; prioritize credentials, paging and error propagation.
2. **IMP-02 / P2:** Generate exports to a temporary sibling file and replace the destination only after success, preserving existing reports on failure.
3. **IMP-03 / P3:** Replace the manual catalog CSV parser with a standard parser and finish localization of direct ViewModel messages. Evaluate any additional dependency separately.

These proposals were not implemented as part of the remediation.

## Session Handoff

Resume from `main` and `PROJECT_STATUS.json`. Offline remediation is complete; the remaining qualification tasks are AUD-07 and AUD-08. Optional improvements IMP-01/02/03 remain proposals. No additional full-suite run is needed unless source changes or new failures justify it. Build/test/security artifacts are ignored by Git and retained locally; regenerate the installer in another environment if needed. The user explicitly authorized commit and push on `main` for this handoff. The final chat records the committed revision, remote equality, and checkout cleanliness.
