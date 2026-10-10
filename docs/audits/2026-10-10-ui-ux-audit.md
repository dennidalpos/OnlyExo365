# UI/UX remediation — 2026-10-10

Historical offline WPF evidence; no dependencies, IPC changes or production tenant operations. Interactive qualification remains UX-08 in [PROJECT_STATUS.json](../../PROJECT_STATUS.json). Prior handoff statements are not authorization to commit/push subsequent work.

## Findings and repairs

| ID | Priority | Repair / recorded evidence |
| --- | --- | --- |
| UX-01 | P1 | DispatcherTimer debounce shared with groups; preserved 300/200 ms delays, removed unused DebounceAsync, disposed callbacks. Dispatcher-bound updates, latest-only execution and disposal pass. |
| UX-02 | P1 | Owned ShowDialog; cancel initially/default focused; Escape/close cancel; manual DispatcherFrame removed. Contact/security/resource/public-folder/group/Mail Flow targets captured before confirmation. Cancellation submits no write; full keyboard journey remains UX-08. |
| UX-03 | P1 | Restored PART_Popup/PART_EditableTextBox and Text bindings on all 12 editable security selectors. Typed/selected values reach save. |
| UX-04 | P1 | Shared action foreground/background pairs across hover/pressed/focus; selected content inherits foreground; visible focus. Contrast changed from primary 2.82 / danger 2.22 to 4.53 / 4.67; hover pairs pass. |
| UX-05 | P2 | Snapshot dirty detection; explicit discard/stay on selection/New/Refresh/navigation/close; restore selection after refusal; block editors during detail loading. Six editor groups covered: mailboxes, groups, resources, public folders, contacts, security; reverting a field clears dirty state. |
| UX-06 | P2 | Editable/search labels and automation names; dynamic system high-contrast palette, reversible to ordinary colors. Catalog/label/palette structure passes; real Narrator/theme checks remain UX-08. |
| UX-07 | P2 | Wrapping Trace filters/footer, compact status selector, resizable/collapsible details; 320-DIP Dashboard/Compliance/Security grids retain recycling. Offscreen bounds/container checks pass; actual DPI/performance remains UX-08. |
| IMP-03 UI portion | P3 | IT/EN client messages, confirmations, summaries, field guidance, filter captions and selected-culture dates/numbers; resource keys/format arguments agree. Technical identifiers/tenant data unchanged; broader proposal remains open. |

Additional repairs: mailbox settings/retention requests now validate and confirm before the first write, including auto-reply; cancelling submits zero writes. Saving a security policy explicitly confirms loss of other drafts; refusal preserves both and submits no write. LocExtension releases dead-target/shutdown-Dispatcher subscriptions.

Sources: [Views](../../src/OnlyExo365.Shell/Views), [ViewModels](../../src/OnlyExo365.Shell/ViewModels), [DarkTheme](../../src/OnlyExo365.Shell/Themes/DarkTheme.xaml), [DebounceHelper](../../src/OnlyExo365.Shell/Helpers/DebounceHelper.cs), [ErrorDialogService](../../src/OnlyExo365.Shell/Services/ErrorDialogService.cs), [ThemePaletteService](../../src/OnlyExo365.Shell/Services/ThemePaletteService.cs), [registrar](../../src/OnlyExo365.Shell/Bootstrap/AppShellModuleRegistrar.cs), [localization](../../src/OnlyExo365.Shell/Localization).

## Recorded verification

Baseline: Debug compiled; 165 targeted presentation/navigation/editor tests passed; preceding full suite had 768. Probes reproduced cross-thread CollectionView updates, hidden editable text and unnamed popup parts.

Final UI revision: Debug 0 errors; **789 passed, 0 failed/skipped**, 21 new regressions. Architecture passed (App 62/80, registrar 161/190); `git diff --check` passed. Commands are in [workflow](../development/workflow.md).

- UiInteractionTests: debounce/disposal, safe dialogs, editable/selected ComboBox behavior, contrast/palette, culture, layout/virtualization, reversible drafts, navigation and cancellation before writes.
- ContactsViewModelTests / MailSecurityViewModelTests: immutable confirmed requests and draft preservation, including save-induced loss of another policy draft.
- UiAccessibilityTests: names/labels, IT/EN keys/arguments and selected-culture numbers.
- Offscreen Trace: 650×420 / 950×620 page viewports, IT/EN, with/without details; visible command/paging buttons stay inside bounds. These approximate page space, not actual window/DPI qualification.
- Virtualization: 1,000 rows; 1–50 realized containers; height ≤320 DIP inside an outer ScrollViewer. This proves the pattern, not each populated page's performance.

Evidence: `artifacts/test-results/onlyexo365-tests.trx` and native Debug build output. Earlier Release/security results belong to the [critical audit](2026-10-10-critical-audit.md); later full-suite/script evidence belongs to the [PowerShell audit](2026-10-10-powershell-audit.md).

## Startup follow-up

`scripts/start.ps1` exit -1 came from MainWindow.Show activating the Language binding. UI log `2026-10-10T13:49:38Z` records BindingExpressionBase.GetCulture failing while converting its own target language to XmlLanguage. Script failure propagation was correct.

MainWindow now sets `ConverterCulture=en-US`; inherited display culture still follows IT/EN. The actual-root-binding offscreen regression reproduced the exception before repair and checks EN/IT/EN date formatting afterward, consistent with [Microsoft ConverterCulture](https://learn.microsoft.com/en-us/dotnet/api/system.windows.data.binding.converterculture).

**31 focused launch/interaction/localization/exception tests passed**, 0 failures/skips; architecture/diff checks passed. Actual `pwsh -NoProfile -File .\start.ps1` from scripts, with EXO disabled only in the child, opened Debug Shell, started Worker 7.6.6, authenticated IPC and closed normally: exit 0, Worker terminated (Shell PID 9932 / Worker PID 14968). Approved EXO was unavailable; no tenant connection. Existing Release publish/setup predate this fix: regenerate before AUD-07.

UX-08, AUD-07/AUD-08 and IMP-03/04 remain separate; closure checks are in [qualification](../development/qualification.md). Sources: [WPF threading](https://learn.microsoft.com/en-us/dotnet/desktop/wpf/advanced/threading-model), [ComboBox parts](https://learn.microsoft.com/en-us/dotnet/desktop/wpf/controls/combobox-styles-and-templates), [owned dialogs](https://learn.microsoft.com/en-us/dotnet/desktop/wpf/windows/how-to-open-window-dialog-box), [accessible text](https://learn.microsoft.com/en-us/windows/apps/design/accessibility/accessible-text-requirements), [virtualization](https://learn.microsoft.com/en-us/dotnet/desktop/wpf/advanced/optimizing-performance-controls).
