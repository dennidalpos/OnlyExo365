# P2 remediation — 26–27 October 2026

Batch: REV-20261026-03, -04, -05, -06, -07 and -09, selected from the authoritative SecondBrain tracker after the P1 batch. The existing P1 working-tree changes were preserved. No commit, push, publication, tenant write or installer execution was requested or performed.

## Changes and acceptance evidence

| Task | Result | Executable evidence |
| --- | --- | --- |
| -03: automatic license catalog | Due checks run before changing LastCheckedUtc; checks that are not due preserve it. An attempt re-arms the one-shot timer after success/failure/cancellation. | LicenseCatalogUpdateServiceTests: controlled HTTP/TimeProvider, first startup, Daily/Monthly expiry, Disabled, forced refresh, failed/cancelled download preserving the catalog, subsequent timer callback. |
| -04: tenant session caches | Shell session generation advances on Worker/Exchange state, user or organization changes. Dashboard and retention policies use generation-scoped keys, clear visible/cached data and reject late results. | SessionCacheRegressionTests: late A response after disconnect/reconnect B or Worker restart, reread before TTL, cached reuse within B, subsequent switch to C. WPF tests run offscreen. |
| -05: failed reads | Connector queries fail if either direction fails; empty successful reads stay successful. Tenant/available/user license queries propagate errors instead of warning-only empty results. The shared error-discarding helper was removed; its mail-security/mobile capability callers now use the checked helper. Informational warnings and structured partial Dashboard/admin results remain supported. | ReadFailureRegressionTests: access denied, missing cmdlet, transient errors, uninitialized engine, empty successful read, informational warning. MailFlowViewModelTests checks error presentation and preservation of existing rows. |
| -06: cloud routing | One JSON map feeds Worker, Shell bootstrap and tenant script: Global/USGov/USGovDoD/China. Graph receives Environment for all four modes; delegated China requires a local ApplicationId. Compliance routing includes ConnectionUri and authority; unsupported DeviceCode/ManagedIdentity are rejected, search-only requires Interactive. | CloudRoutingRegressionTests executes real builders and extracted tenant functions against local cmdlets for all four clouds and authentication modes, checks bootstrap syntax, unknown clouds and unsupported combinations. Native EXO 3.10.1 Connect-IPPSSession signature inspected without connecting. |
| -07: rule matcher | Typed per-rule Match/NoMatch/Indeterminate results, explicit reason strings and localized IT/EN limitations. Disabled rules cannot match; Audit, unresolved membership/identities, unavailable metadata and unsupported predicates remain indeterminate. Subject comparisons use literal case-insensitive text. | RuleMatcherRegressionTests: state, Audit, membership, unsupported/missing metadata, unresolved identity, exceptions, subject wildcard characters and failed reads. GeneratedPowerShellExecutionTests tests exception behavior; MailFlowViewModelTests verifies displayed reasons and limitations. |
| -09: IPC send deadline | Starts the budget before the send lock; linked token reaches write, flush and response wait. Notification of cancellation has a separate two-second best-effort budget and does not delay the caller. No write retry added. | IpcSendDeadlineRegressionTests: held lock, blocked controlled stream and delayed response, each with timeout/manual cancellation, bounded completion and empty pending/event registration dictionaries. Existing real named-pipe regression tests also pass. |

## Verification

Windows x64, exact SDK 10.0.401 from the retained `%TEMP%/OnlyExo365-nuget-20261026/dotnet` root, with process-local PATH/DOTNET_ROOT. `global.json`, package references and lockfiles retain their existing versions. The SDK's selected installed runtime is 10.0.14. Existing ignored build/evidence artifacts and retained NuGet scratch were preserved.

| Command/check | Outcome |
| --- | --- |
| `dotnet test -c Release --artifacts-path artifacts/build --filter 'FullyQualifiedName~RegressionTests\|FullyQualifiedName~LicenseCatalog\|FullyQualifiedName~MailFlow\|FullyQualifiedName~GeneratedPowerShellExecutionTests\|FullyQualifiedName~WorkerCommandTests\|FullyQualifiedName~PowerShellScriptBehaviorTests\|FullyQualifiedName~InteractiveExchangeBootstrapServiceTests' --nologo` | 265 passed, zero failed/skipped |
| `dotnet test -c Release --artifacts-path artifacts/build --nologo` | 887 passed, zero failed/skipped; 44 added since the P1 batch |
| Final retention-policy state adjustment | 4 focused session-cache/editor tests passed, followed by the full suite again; final TRX retained |
| `pwsh -NoProfile -File scripts/build.ps1 -Configuration Release -RuntimeIdentifier win-x64 -Clean:$false -LockedMode` | Passed, zero errors; clean disabled to preserve prior outputs |
| `pwsh -NoProfile -File build/assert-architecture-constraints.ps1` | Passed |
| Native PowerShell Parser.ParseFile over build/scripts | 32 scripts passed |
| Installed approved Graph Authentication 2.41.1 | Connect-MgGraph exposes Environment; Get-MgEnvironment returns the four mapped clouds/endpoints without sign-in |
| `git diff --check` | Passed |

TRX and logs remain in `artifacts/test-results/p2-20261026/`; authoritative completion records, per-criterion results, source/evidence SHA-256 hashes and original tracker records are in the SecondBrain dossier under `history/2026-10-27-p2-remediation/`. SecondBrain validation/Sync/Check results are retained there as well.

## Confirmed separate residue

Worker Compliance app-certificate authentication with only CertificateSubjectName still generates `Connect-IPPSSession -CertificateSubjectName`, but the installed exact EXO 3.10.1 cmdlet does not expose that parameter. CertificateThumbprint is exposed. The tenant validation script already resolves a unique subject to a thumbprint; the Worker builder does not. This pre-existing defect is recorded as REV-20261027-01, outside the six-task batch. No certificate-store change or authentication attempt was performed. Documentation points operators to the explicit-thumbprint path pending remediation.

## Sources and limits

- [Microsoft Graph authentication/environment guidance](https://learn.microsoft.com/en-us/powershell/microsoftgraph/authentication-commands?view=graph-powershell-1.0) supports the four selected Graph environments and the China custom-app constraint.
- [Security & Compliance connection guidance](https://learn.microsoft.com/en-us/powershell/exchange/connect-to-scc-powershell?view=exchange-ps) documents national-cloud endpoints and delegated search requirements.
- [Connect-IPPSSession reference](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/connect-ippssession?view=exchange-ps), checked against the installed EXO 3.10.1 command signature for the parameters noted above.

Current documentation alone does not qualify every pinned-module/cloud combination. No real tenant authentication, module network call, full Exchange rule-engine simulation, installer, accessibility, DPI or populated-grid performance was qualified. Unknown predicate metadata intentionally yields indeterminate. Controlled-stream send tests do not qualify recovery from a partially transmitted frame on a saturated real pipe. UX-08 and CLEAN-01 remain blocked; the newly confirmed certificate residue is not silently included in this batch.
