# Checkbox & Filter Logic Analysis

Questo documento aggiorna la mappatura **end-to-end** di checkbox e filtri UI:
**View → ViewModel → Request DTO → Worker/PowerShell → refresh UI**.

## 1) Mailboxes / Shared Mailboxes

### Search (`SearchQuery`)
- **UI**: textbox ricerca.
- **ViewModel**: debounce + refresh pagina.
- **DTO**: `GetMailboxesRequest.SearchQuery`.
- **Worker**: filtro su `DisplayName`, `PrimarySmtpAddress`, `Alias`.

### Recipient Type (`RecipientTypeFilter`)
- **UI**: combo tipo mailbox.
- **ViewModel**: refresh immediato alla modifica.
- **DTO**: `GetMailboxesRequest.RecipientTypeDetails`.
- **Worker**: filtro PowerShell su `RecipientTypeDetails`.

## 2) Deleted Mailboxes

### UPN query (`UpnQuery`)
- **UI**: campo verifica UPN.
- **ViewModel**: valorizza `_activeSearchQuery` e lancia refresh.
- **DTO**: `GetDeletedMailboxesRequest.SearchQuery`.
- **Worker**: filtro su `DisplayName`, `PrimarySmtpAddress`, `UserPrincipalName`, `Alias`.

## 3) Distribution Lists

### Include Dynamic (`IncludeDynamic`)
- **UI**: checkbox “Includi dinamiche”.
- **ViewModel**: toggle → refresh elenco.
- **DTO**: `GetDistributionListsRequest.IncludeDynamic`.

### Allow External Senders (`AllowExternalSenders`)
- **UI**: checkbox in dettaglio DL.
- **ViewModel**: modifica staged fino a Save.
- **DTO**: mapping verso `RequireSenderAuthenticationEnabled` invertito.

## 4) Mail Flow (stato aggiornato)

### Rule Enabled (`RuleEnabled`)
- **UI**: checkbox “Enabled” sezione Rule.
- **Abilitazione edit**: `CanEditSelectedRule` (`SelectedRule != null && !IsLoading`).
- **DTO**: `UpsertTransportRuleRequest.Enabled`.
- **Persistenza**: Save Rule → worker `UpsertTransportRule` → refresh liste.

### Connector Enabled (`ConnectorEnabled`)
- **UI**: checkbox “Enabled” sezione Connector.
- **Abilitazione edit**: `CanEditSelectedConnector`.
- **DTO**: `UpsertConnectorRequest.Enabled`.
- **Persistenza**: Save Connector → worker `UpsertConnector` → refresh liste.
- **Conferma rischio**: se passa da Enabled→Disabled su connector selezionato, richiesta conferma tenant-wide.

### Domain Make Default (`DomainMakeDefault`)
- **UI**: checkbox “Make Default”.
- **Abilitazione edit**: `CanEditSelectedDomain`.
- **DTO**: `UpsertAcceptedDomainRequest.MakeDefault`.
- **Persistenza**: Save Domain → worker `UpsertAcceptedDomain` → refresh liste.

### Filtri test rule
- **UI**: `TestSender`, `TestRecipient`, `TestSubject`.
- **Validazione**: sender/recipient email valide.
- **DTO**: `TestTransportRuleRequest`.

## 5) Message Trace

### Sender/Recipient/Date filters
- **DTO**: `GetMessageTraceRequest`.
- **Status filter** (`All/Delivered/Failed/Pending`): client-side su `AllMessages`.
- **Dettagli**: `GetMessageTraceDetailsRequest` su elemento selezionato.

### Export
- Export aggiornato da CSV a **Excel `.xlsx` formattato**.
- L’export applica intestazione formattata e righe dati filtrate correnti.

## 6) Logs

### Text filter (`SearchFilter`)
- case-insensitive su source + message.

### Level filter (`FilterLevel`)
- soglia minima severità (`entry.Level >= FilterLevel`).

## Regole operative trasversali

- Le checkbox di dettaglio sono modifiche **staged** fino a Save.
- La disabilitazione avviene solo quando esplicitamente consentita da stato pagina (`!IsLoading` + selezione valida).
- Errori di persistenza mostrano messaggio user-friendly e vengono loggati con riferimento diagnostico.
