# Documentation/source review — 2026-10-10

Reviewed all 13 existing Markdown documents, PROJECT_STATUS.json and source comments across C#, PowerShell, XAML, resources, build/installer and CI files. Pre-existing edits were preserved; audit/runtime fixes were not reimplemented. No dependencies, generated catalogs, lockfiles or executable behavior changed.

## Reconciliation

| Previous statement | Source-confirmed correction |
| --- | --- |
| Weekly catalog mode / Weekly sample | LicenseCatalogConfiguration supports Disabled, Daily, Monthly; shipped mode is Daily |
| Environment variables apply to every configuration section | Exchange has environment overlays; catalog uses file overlays and persisted metadata mode |
| All user state is LocalAppData | Language preferences use roaming AppData/OnlyExo365/OnlyExo365 |
| Log retention is 30 days | PersistentLogWriter defaults to 14 days |
| Catalog paths expand %LocalAppData% | Custom LocalCachePath is used literally; null selects the default |
| Session token is DPAPI-protected | Supervisor passes the session token directly in the child environment; DPAPI protects separate password references |
| Fixed pipe names / immediate disconnect on failed handshake | Names include user hash/session suffix; domain requests remain blocked after failed handshake |
| Console toggle uses handshake | Dedicated SetWorkerConsoleVisibility request |
| MaxMessageSizeBytes enforces UTF-8 bytes | Callers pass string.Length after full-line allocation; byte enforcement/bounded reads remain IPC-01 |
| Conceptual layers are enforced project boundaries | Existing projects contain responsibility folders/types; no separate layer assemblies |
| Local gate includes live/interactive checks | Gate packages offline; installer smoke, tenant and UI qualification are separate |
| Smoke covers only startup | Default smoke also installs/reinstalls/uninstalls; SkipRealInstall omits that qualification |
| License writes are always opt-in defaults | Contract default scopes are empty; shipped appsettings includes LicenseAssignment.ReadWrite.All |
| PowerShell remediation package is current | Release artifacts predate the later MainWindow binding fix; regenerate before AUD-07 |

Sources are linked beside each domain guide. Requirements/commands live in workflow; storage/authentication/diagnostics/signing/qualification have separate guides. Removed the dated SDK ADR after consolidating its runtime rationale in architecture. Audit tables retain finding IDs, reproductions, verification counts, evidence paths and limitations; obsolete line references and conversational handoffs were removed. CLEAN-01 paths remain in the retained-fixtures record. README is the navigation index; tracker owns open work.

## Comment policy and verification

Kept reasons, ownership, threading, fallback and atomicity constraints; removed labels or narration already expressed by symbols/bindings and shortened XML summaries. Compiler directives, Requires, production strings and generated PowerShell comments remain intact.

Baseline/final architecture and whitespace checks passed. Native Roslyn comparison covers all **396 C# files**: executable tokens and directive tokens unchanged, comment blocks **96 → 55**. Native PowerShell parser comparison covers **32 scripts**, zero parse errors and unchanged non-comment tokens. XML/resource contents, local links, source pins, tracker IDs, encoding and line endings are checked separately. The Roslyn verifier runs in a disposable SDK console because loading SDK Roslyn into pwsh conflicted with its existing assembly version.

IPC-01 / P2: a local native Contracts probe confirmed that a JSON string with **4,000,002 characters / 12,000,002 UTF-8 bytes** passes the guard named MaxMessageSizeBytes (10,485,760). IpcClient/IpcServer pass string.Length and check after ReadLineAsync completes; the read-buffer constant is unused. Byte-accurate size enforcement and bounded reading remain open; no protocol implementation was changed in this documentation pass. Existing oversize tests are ASCII-only and do not cover the multibyte case.

The concrete IPC discrepancy prompted these checks:

```powershell
pwsh -NoProfile ./scripts/agents/compile.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBootstrap
dotnet test tests/OnlyExo365.Tests/OnlyExo365.Tests.csproj -c Debug --no-build --no-restore --artifacts-path artifacts/build --filter 'FullyQualifiedName~OnlyExo365.Tests.IpcServerOversizeTests|FullyQualifiedName~OnlyExo365.Tests.IpcContractsTests' --logger 'trx;LogFileName=documentation-ipc.trx' --results-directory artifacts/documentation-review-tests
pwsh -NoProfile ./build/assert-architecture-constraints.ps1
git diff --check
```

Debug compilation passed with **0 errors**; focused IPC tests: **20 passed, 0 failed/skipped**, 203 ms. Passing existing tests does not close IPC-01. Full suite, packaging, tenant and interactive UI were not rerun. Earlier full-suite numbers remain historical evidence. AUD-07, AUD-08, UX-08 and CLEAN-01 remain open alongside IPC-01.

Final verification: **114 valid local links**, **35 unchanged XML/resource documents**, UTF-8 without BOM and existing line endings preserved; no obsolete ADR references. Retained local evidence: `artifacts/documentation-review-20261010-results.json` and `artifacts/documentation-review-tests/documentation-ipc.trx`. Earlier CLEAN-01 fixtures remain preserved.

Verification scratch cleanup was rejected by automatic approval review: `blocked by policy`. The requested action was literal recursive removal of this session's `D:\GITHUB\OnlyExo365\artifacts\documentation-review-20261010`. Preflight confirmed repository/artifacts containment, Git ignore, no tracked files or reparse points and 90 owned entries. No narrower reason was supplied; deletion was not retried. The directory retains comparison baselines, disposable Roslyn verifier/build outputs and the local IPC probe, with no product/personal state. Cleanup remains a CLEAN-01 residual in the [fixture record](2026-10-10-retained-fixtures.md).

Behavioral primary sources refreshed on 2026-10-10: [Microsoft Compliance/eDiscovery](https://learn.microsoft.com/en-us/powershell/exchange/connect-to-scc-powershell?view=exchange-ps), [Graph throttling](https://learn.microsoft.com/en-us/graph/throttling), [Trace V2](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-messagetracev2), [SKU reference](https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference), [WPF ConverterCulture](https://learn.microsoft.com/en-us/dotnet/api/system.windows.data.binding.converterculture), [Inno uninstall](https://jrsoftware.org/ishelp/topic_uninstalldeletesection.htm), [Inno signing](https://jrsoftware.org/ishelp/topic_setup_signtool.htm). Version pins were checked against source, not inferred to be latest releases.
