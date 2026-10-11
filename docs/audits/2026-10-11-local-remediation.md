# Local remediation — 2026-10-11

This log records task closure against [PROJECT_STATUS.json](../../PROJECT_STATUS.json). Installer and tenant qualification were removed from the task list at user request; their quality gates remain unverified.

## IPC-01 — UTF-8 byte limit and bounded line reads

The shared `IpcLineReader` reads raw pipe bytes in 256 KiB chunks and caps each JSON line at 10,485,760 bytes. Oversized lines are discarded through the next newline without allocating the full line. The client and worker use it for request, response, event and handshake reads. Outbound guards use UTF-8 byte counts. The JSON-line protocol, timeout/cancellation envelopes and write-retry behavior are unchanged.

Baseline: 24 focused IPC tests passed. After repair: Debug compile succeeded with 0 errors; 27 focused IPC tests passed, including exact-limit multibyte, over-limit discard/recovery, BOM and CRLF cases. Architecture constraints and `git diff --check` passed. No full suite was run.

Primary reference: [Microsoft .NET Encoding.GetByteCount](https://learn.microsoft.com/en-us/dotnet/api/system.text.encoding.getbytecount?view=net-10.0).

## IMP-02 — atomic exports

XLSX reports and diagnostic ZIP archives now use the shared `AtomicFileWriter`. Each output is completed in a unique temporary file in the destination directory before it replaces an existing report. A generation failure preserves the previous report and removes the partial file. Source log retention still follows the configured policy.

Baseline: 9 focused export/log tests passed. After repair: Debug compile succeeded with 0 errors; 11 focused tests passed, including a failing XLSX row producer, temporary-file cleanup, existing-file replacement and ZIP overwrite. Architecture constraints and `git diff --check` passed. No full suite was run.

Primary reference: [Microsoft .NET File.Replace](https://learn.microsoft.com/en-us/dotnet/api/system.io.file.replace?view=net-10.0).

## IMP-01 — generated PowerShell execution coverage

Five offline runspace cases now execute generated transport-rule update/preview, remote-domain create/configure and address-list update scripts against local dummy cmdlets. They verify named parameter binding, escaped identity, create-before-configure order, predicate exceptions and command failure propagation. This extends the earlier repaired-path coverage without tenant access or writes. It does not qualify all generated commands or real Exchange behavior; that residual is recorded in the tracker risks.

Baseline: 4 relevant command-composition tests passed. After the new cases: Debug compile succeeded with 0 errors; 18 focused execution/regression tests passed, including the five new cases. Architecture constraints and `git diff --check` passed. No full suite was run.

Primary reference: [Microsoft PowerShell splatting](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_splatting).

## IMP-03 — catalog CSV parsing

The Shell catalog downloader now uses .NET `TextFieldParser` in place of the manual line parser. It accepts quoted commas, doubled quotes, embedded newlines, blank lines, whitespace and a UTF-8 BOM, and reports malformed rows as `CatalogDownloadException`. The existing catalog grouping and atomic cache write path are unchanged. No package was added.

Baseline: 12 downloader tests passed. After replacement: Debug compile succeeded with 0 errors; 19 focused downloader/file-store tests passed, including multiline and malformed CSV. Architecture constraints and `git diff --check` passed. No full suite was run.

Primary reference: [Microsoft .NET TextFieldParser.ReadFields](https://learn.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.fileio.textfieldparser.readfields?view=net-10.0).

Localization evaluation: The user chose to keep Worker and tenant technical messages in their original wording for support. LOC-01 is resolved by that decision; IPC and source diagnostics are unchanged. A remaining Shell-owned audit queue presentation gap was fixed: task state, title, filter labels, summary and displayed request time now follow the selected IT/EN locale, including a live language change. Raw `ProgressEventPayload.StatusMessage`, Compliance hold warnings, errors and tenant data remain unchanged. Eight focused Compliance/accessibility tests passed before this correction and nine passed afterward; Debug test-project compilation and resource-key checks passed.

Primary reference for live label updates: [Microsoft WPF weak event patterns](https://learn.microsoft.com/en-us/dotnet/desktop/wpf/events/weak-event-patterns).

## IMP-04 — eager page construction measurement

`AppBootstrapCompositionTests` creates the 18 page ViewModels with an offline Worker stub. Two local Debug runs of its four existing cases passed (4/4 each). The first case took 157 ms and 152 ms; subsequent cases took 4–21 ms. These per-case durations include JIT, fixture setup and assertions, so they are a proxy upper bound for warm composition rather than an isolated constructor or full UI startup benchmark. No lazy-page change is justified by this evidence. Navigation and page lifetime remain unchanged. Actual window startup, populated data and responsiveness remain in UX-08.

Primary reference: [Microsoft WPF application startup guidance](https://learn.microsoft.com/en-us/dotnet/desktop/wpf/advanced/application-startup-time).
