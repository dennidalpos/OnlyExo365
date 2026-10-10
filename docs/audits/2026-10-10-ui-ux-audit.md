# UI and UX audit — 2026-10-10

The approved WPF repairs are implemented and verified offline. Interactive qualification remains open as UX-08 in [PROJECT_STATUS.json](../../PROJECT_STATUS.json). No new dependencies, IPC changes or production tenant operations were performed. The user subsequently authorized commit and push on main for session handoff.

## Baseline

Debug compilation passed with 0 errors. The targeted presentation/navigation/editor baseline passed 165 tests; the prior complete suite contained 768 tests. Offline WPF probes reproduced cross-thread CollectionView updates and an editable ComboBox whose text part remained hidden and whose popup part was unnamed. Button contrast was 2.82:1 (primary) and 2.22:1 (danger).

## Findings, evidence and closure criteria

| ID | Priority | Evidence and repair | Closure criterion | Outcome |
| --- | --- | --- | --- | --- |
| UX-01 | P1 | [DebounceHelper](../../src/OnlyExo365.Shell/Helpers/DebounceHelper.cs): replace thread-pool scheduling with DispatcherTimer; share it with groups; preserve 300/200 ms delays; remove unused DebounceAsync; release callbacks at shutdown | Bound CollectionView updates run on its Dispatcher; only the latest callback executes; disposal cancels pending work | Verified offline |
| UX-02 | P1 | [ErrorDialogService](../../src/OnlyExo365.Shell/Services/ErrorDialogService.cs), [ConfirmationDialog](../../src/OnlyExo365.Shell/Views/ConfirmationDialog.xaml): Owner + ShowDialog; cancel gets initial/default focus; Escape/close cancel; remove manual DispatcherFrame. Capture contact, security policy, resource, public folder, group settings and Mail Flow deletion requests before confirmation | Cancellation submits no write; submitted target and payload match the captured request; safe cancel/close behavior | Verified offline; full keyboard journey in UX-08 |
| UX-03 | P1 | [DarkTheme](../../src/OnlyExo365.Shell/Themes/DarkTheme.xaml), [MailSecurityView](../../src/OnlyExo365.Shell/Views/MailSecurityView.xaml): restore PART_Popup/PART_EditableTextBox and bind all twelve editable selectors through Text | Typing and choosing an option update the property used by save; popup and visible editor exist | Verified offline |
| UX-04 | P1 | [DarkTheme](../../src/OnlyExo365.Shell/Themes/DarkTheme.xaml): central action text/background pairs; hover/pressed states retain the pair; selected item content inherits selected foreground; visible keyboard focus | Normal action text reaches 4.5:1 in default/hover/focus states | Verified offline: primary 4.53:1, danger 4.67:1; hover pairs also pass |
| UX-05 | P2 | [Contacts](../../src/OnlyExo365.Shell/ViewModels/ContactsViewModel.cs), [MailSecurity](../../src/OnlyExo365.Shell/ViewModels/MailSecurityViewModel.cs), [PublicFolders](../../src/OnlyExo365.Shell/ViewModels/PublicFoldersViewModel.cs), [registrar](../../src/OnlyExo365.Shell/Bootstrap/AppShellModuleRegistrar.cs): compare snapshots; guard destructive replacement; restore selection after refusal; discard before approved navigation; guard window close; block editing during detail loads | Cancelled selection/New/Refresh/navigation retains draft and selection; reverting a field clears dirty state; no implicit discard of another security editor on save | Verified offline across the six registered editor groups: mailboxes, groups, resources, public folders, contacts, security |
| UX-06 | P2 | All [Views](../../src/OnlyExo365.Shell/Views), [ThemePaletteService](../../src/OnlyExo365.Shell/Services/ThemePaletteService.cs): LabeledBy/name for editable fields and searches; dynamic resources switch to system high-contrast brushes | Inputs have accessible labels; high-contrast palette restores ordinary colors when disabled | Catalog/label structure and palette mapping verified offline; Narrator and real Windows themes in UX-08 |
| UX-07 | P2 | [MessageTraceView](../../src/OnlyExo365.Shell/Views/MessageTraceView.xaml), [MainWindow](../../src/OnlyExo365.Shell/Views/MainWindow.xaml): wrapping filters/footer, compact status selector, resizable/collapsible detail panel. Bound Dashboard/Compliance/Security grids to 320 DIP and retain recycling | Commands/paging stay inside compact viewports; large grids realize viewport rows instead of the full collection | Verified offline; actual DPI/window matrix and performance profiling in UX-08 |
| IMP-03 | P3 | [catalogs](../../src/OnlyExo365.Shell/Localization/Strings.resx), [LocalizationService](../../src/OnlyExo365.Shell/Localization/LocalizationService.cs), [LocalizedFormatConverter](../../src/OnlyExo365.Shell/Localization/LocalizedFormatConverter.cs): IT/EN messages, confirmations, summaries, required-field guidance, localized filter captions, selected-culture dates/numbers, descriptive resource keys | Catalog keys/format arguments match; date and number formatting follow selected locale; disabled Save explains requirements | Targeted UI portion verified offline; broader IMP-03 remains a separate proposal |

## Additional defects corrected

- Mailbox settings and retention writes previously ran before confirmation. [MailboxDetailsViewModel](../../src/OnlyExo365.Shell/ViewModels/MailboxDetailsViewModel.cs) now builds and validates all requests, then confirms before the first write. Cancelling submits zero settings/retention/auto-reply writes.
- Saving one security policy reloads every policy editor. When other drafts exist, the single confirmation explicitly offers saving and discarding the other changes, or staying. Refusal preserves both editors and submits no request.
- A selected ComboBoxItem previously retained its inner text's ordinary foreground. The text now inherits the selected foreground.
- Localization notifications could target a closed Dispatcher. [LocExtension](../../src/OnlyExo365.Shell/Localization/LocExtension.cs) releases subscriptions for dead targets or shutting-down Dispatchers.

## Verification

Executed against the final source on 2026-10-10:

```powershell
pwsh ./scripts/agents/compile.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBootstrap
pwsh ./scripts/agents/test.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBootstrap -NoBuild -NoRestore
pwsh ./build/assert-architecture-constraints.ps1
git diff --check
```

- Debug compilation: **0 errors**. Full suite: **789 passed, 0 failed, 0 skipped** (21 additional regressions).
- Architecture constraints: passed; App 62/80 and registrar 161/190 non-empty lines. Diff whitespace check: passed.
- [UiInteractionTests](../../tests/OnlyExo365.Tests/UiInteractionTests.cs): Dispatcher-bound collection updates; latest-only scheduling and disposal; safe dialog action and close; typed/selected ComboBox value; rendered selected foreground; contrast and system palette; inherited date culture; bounded grid realization; compact Trace layout; reversible dirty state; navigation preservation/discard; cancellation before mailbox writes.
- [ContactsViewModelTests](../../tests/OnlyExo365.Tests/ContactsViewModelTests.cs) and [MailSecurityViewModelTests](../../tests/OnlyExo365.Tests/MailSecurityViewModelTests.cs): immutable confirmed requests and cancelled draft replacement, including save-induced loss of another policy draft.
- [UiAccessibilityTests](../../tests/OnlyExo365.Tests/UiAccessibilityTests.cs): editable input names/labels, localization key presence, IT/EN parity and format arguments, selected-culture numbers.
- Offscreen Trace layout: page viewports **650×420** and **950×620**, IT/EN, with/without details. These approximate the page space in 900×600 and 1200×800 windows; every visible command/paging button remained within bounds.
- Virtualization probe: **1,000 rows**, realized containers **between 1 and 50**, bounded height **≤320 DIP** inside an outer ScrollViewer. This verifies the container pattern; production performance and each populated page have not been profiled.
- Evidence: `artifacts/test-results/onlyexo365-tests.trx` and the native build output are retained locally, gitignored. Earlier Release/security/packaging results belong to the critical audit and were not rerun for these UI changes.

## Sources

Microsoft documentation consulted on 2026-10-10: [WPF threading](https://learn.microsoft.com/en-us/dotnet/desktop/wpf/advanced/threading-model), [ComboBox parts](https://learn.microsoft.com/en-us/dotnet/desktop/wpf/controls/combobox), [owned modal windows](https://learn.microsoft.com/en-us/dotnet/desktop/wpf/windows/how-to-open-window-dialog-box), [accessible text](https://learn.microsoft.com/en-us/windows/apps/design/accessibility/accessible-text-requirements), [control virtualization](https://learn.microsoft.com/en-us/dotnet/desktop/wpf/advanced/optimizing-performance-controls).

## Remaining qualification and proposals

- **UX-08 / P2:** interactive windows at 900×600 and 1200×800; DPI 100/150/200%; IT/EN; full keyboard journey, Escape, owner modality and Narrator; actual Windows high-contrast toggles. These checks are **not verified**: this session provided offline WPF probes, without access to native desktop automation. Inspect all repaired pages with realistic long text and data, and measure actual realized rows and responsiveness.
- **IMP-03 / P3:** the standard CSV parser proposal and wider localization of technical operation descriptions/worker-provided status messages remain separate. This repair translates client presentation text without changing protocol identifiers or tenant-provided data.
- Lazy creation of the 18 startup pages remains a measurement-led proposal. Measure startup cost before changing page lifetime.
- **AUD-07/AUD-08:** isolated installer and dedicated test-tenant qualification remain open. Regenerate the Release package before qualifying this UI revision. Offline regressions do not qualify actual assistive technology or Microsoft 365 behavior.

## Session handoff

The user requested saving the completed work, committing and pushing on `main`, and changing session. This audit, the source repairs and the tracker are included in the same handoff commit. The final chat records the committed revision and the verified remote/checkout state. Resume from `PROJECT_STATUS.json`; UX-08, AUD-07 and AUD-08 remain open, with IMP-01/02/03/04 tracked separately. Preserve local ignored evidence; regenerate Release artifacts before installer qualification.
