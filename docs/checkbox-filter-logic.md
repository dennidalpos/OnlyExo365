# Checkbox & Filter Logic Analysis

This document maps each user-facing checkbox/filter to its ViewModel logic and backend request fields.

## 1) Mailboxes / Shared Mailboxes

### Search box (`SearchQuery`)
- **UI**: `MailboxListView` / `SharedMailboxListView` text input.
- **Behavior**: updates `SearchQuery` with `UpdateSourceTrigger=PropertyChanged`.
- **Logic**: debounced refresh (`DebounceHelper`, 300ms).
- **Backend mapping**: `GetMailboxesRequest.SearchQuery`.

### Recipient type filter (`RecipientTypeFilter`)
- **UI**: combo box (Mailbox/Shared/Room/Equipment depending on view).
- **Behavior**: immediate refresh when selection changes.
- **Backend mapping**: `GetMailboxesRequest.RecipientTypeDetails`.

## 2) Deleted Mailboxes

### UPN query filter (`UpnQuery`)
- **UI**: dedicated search box.
- **Behavior**: explicit search action populates `_activeSearchQuery`.
- **Backend mapping**: `GetDeletedMailboxesRequest.SearchQuery`.

## 3) Distribution Lists

### Search box (`SearchQuery`)
- **UI**: text input.
- **Behavior**: debounced refresh (300ms via `DispatcherTimer`).
- **Backend mapping**: `GetDistributionListsRequest.SearchQuery`.

### Checkbox: Include dynamic groups (`IncludeDynamic`)
- **UI**: checkbox "Includi dinamiche".
- **Behavior**: refresh is triggered when toggled.
- **Implementation note**: refresh now uses a safe async wrapper that logs exceptions.
- **Backend mapping**: `GetDistributionListsRequest.IncludeDynamic`.

### Checkbox: Allow external senders (`AllowExternalSenders`)
- **UI**: checkbox in group details.
- **Behavior**: marks pending settings changes; does not auto-save.
- **Backend mapping**: converted to inverse `RequireSenderAuthenticationEnabled` when saving.

### Sender allow/deny filters
- **UI**: add/remove lists for:
  - `AcceptMessagesOnlyFrom`
  - `RejectMessagesFrom`
- **Behavior**: in-memory edits tracked against normalized original values.
- **Backend mapping**: `SetDistributionListSettingsRequest.AcceptMessagesOnlyFrom` / `RejectMessagesFrom`.

## 4) Message Trace

### Sender/Recipient filters
- **UI**: text inputs.
- **Behavior**: applied on search.
- **Backend mapping**: `GetMessageTraceRequest.SenderAddress` / `RecipientAddress`.

### Date range filters (`StartDate`, `EndDate`)
- **UI**: date selectors.
- **Behavior**: applied on search; `EndDate` is expanded to end-of-day (`23:59:59`).
- **Backend mapping**: `GetMessageTraceRequest.StartDate` / `EndDate`.

### Page size filter (`PageSize`)
- **UI**: numeric/selection control.
- **Behavior**: affects per-page retrieval and `HasMore` paging logic.
- **Backend mapping**: `GetMessageTraceRequest.PageSize`.

## 5) Logs

### Search filter (`SearchFilter`)
- **UI**: text input.
- **Behavior**: triggers cache invalidation + re-filtering.
- **Filter rule**: case-insensitive match on source/message text.

### Level filter (`FilterLevel`)
- **UI**: combo (All, Debug+, Info+, Warning+, Error).
- **Behavior**: minimum-severity threshold.
- **Filter rule**: keep entries where `entry.Level >= FilterLevel`.

## Operational Notes

- Most filters are **client-driven inputs** that map directly to request DTO fields.
- Most checkboxes in details pages represent **staged changes** and are only persisted when user clicks **Save**.
- List-page filters favor responsiveness (debounce + incremental paging).
