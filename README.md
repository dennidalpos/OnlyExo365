# OnlyExo365 / ExchangeAdmin

Applicazione desktop **WPF** per amministrazione operativa di **Exchange Online**, organizzata a layer e con worker PowerShell separato.

## Scopo

L’applicazione concentra in una sola interfaccia le attività amministrative quotidiane:

- connessione a Exchange Online;
- analisi mailbox (utente, shared, cancellate);
- gestione distribution list (statiche e dinamiche);
- consultazione message trace;
- visibilità centralizzata dei log operativi.

L’obiettivo principale è separare chiaramente UI e logica di esecuzione PowerShell, così da ridurre i blocchi in interfaccia e migliorare la resilienza in caso di errori runtime.

## Architettura del repository

- `src/ExchangeAdmin.Presentation`
  UI WPF, ViewModel, validazioni input, comandi utente e logiche di filtro lato client.
- `src/ExchangeAdmin.Application`
  Use case applicativi e orchestrazione verso i servizi worker.
- `src/ExchangeAdmin.Infrastructure`
  IPC client/server locale e supervisione ciclo vita worker.
- `src/ExchangeAdmin.Worker`
  Esecuzione PowerShell (runspace), import moduli, cmdlet EXO/Graph.
- `src/ExchangeAdmin.Contracts`
  DTO e messaggi IPC condivisi tra Presentation e Worker.
- `src/ExchangeAdmin.Domain`
  Tipologie dominio, gestione errori normalizzati, resilienza.
- `build/`
  Script di build e pulizia.
- `installer/`
  Sorgenti WiX per MSI.

## Flusso operativo principale

1. La UI invia una richiesta tipizzata al layer Application.
2. Infrastructure inoltra la richiesta al worker via IPC locale.
3. Il worker esegue i cmdlet nel runspace PowerShell.
4. I risultati vengono normalizzati (payload/errori/progress) e ritornati alla UI.
5. La UI aggiorna stato, liste, dettagli e log.

## Processi di import/export

### Import moduli PowerShell (worker)

- Il worker importa `ExchangeOnlineManagement` nel runspace dedicato.
- Prima dell’import, viene verificata la disponibilità del modulo in `PSModulePath`.
- In caso di modulo assente o import fallito, l’errore viene normalizzato e propagato ai layer superiori.
- L’import può essere rieseguito durante recovery del runspace.

### Export message trace (UI)

- Il comando export agisce sulla collezione filtrata corrente (`Messages`).
- Formato output: `.xlsx` (OpenXML).
- Header esportati: `Received`, `Sender`, `Recipient`, `Subject`, `Status`, `Size`, `MessageId`, `MessageTraceId`.
- Directory di export:
  - priorità alla variabile ambiente `EXCHANGEADMIN_EXPORT_DIR`;
  - fallback: `%LOCALAPPDATA%\OnlyExo365\exports`.
- Il processo gestisce eccezioni con messaggio utente (`ErrorMessage`) e log applicativo.

### Cartelle artefatti import/export (script)

- Lo script `build/build.ps1` prepara sia cartella `exports` che `imports` (personalizzabili via parametro).
- Lo script `build/clean.ps1` può pulire in modo selettivo `exports` e `imports`, oltre a `artifacts`, `bin/obj`, file temporanei e cache opzionali.

## Logica checkbox e filtri

### Distribution Lists

- **Checkbox `IncludeDynamic`**: include/esclude i gruppi dinamici dalle query lista.
- **Ricerca testuale**: applicata con debounce, invio query normalizzata (trim/null).
- **Checkbox impostazioni lista**:
  - `AllowExternalSenders` aggiorna `RequireSenderAuthenticationEnabled` in modo inverso;
  - variazioni su mittenti ammessi/bloccati attivano `HasPendingSettingsChanges` tramite confronto normalizzato (trim, distinct case-insensitive, ordinamento).
- **Comandi di salvataggio/scarto**: abilitati/disabilitati in base allo stato di modifica pendente.

### Message Trace

- **Filtro stato (`StatusFilter`)**: ammessi solo `All`, `Delivered`, `Failed`, `Pending`.
- Valori non previsti vengono normalizzati ad `All`.
- Il filtro viene applicato localmente su `AllMessages` per aggiornare `Messages` senza roundtrip al worker.
- L’abilitazione dei comandi viene ricalcolata ad ogni applicazione filtro.

### Mailbox List

- **Filtro `RecipientTypeFilter`**: normalizzato a `UserMailbox`/`SharedMailbox`.
- **Ricerca testuale**: debounce lato UI e invio query trim/null.
- **Paginazione**: `LoadMore` basato su `HasMore` + `Skip` corrente.

### Logs

- **Filtro livello minimo (`FilterLevel`)**.
- **Filtro testo (`SearchFilter`)** con normalizzazione trim/case-insensitive.
- Cache locale dei risultati filtrati, invalidata su cambi filtro o su nuovi log.
- Refresh dei risultati con debounce/throttling per ridurre aggiornamenti UI eccessivi.

## Prerequisiti

- Windows (WPF).
- PowerShell 7+ (`pwsh`).
- .NET SDK 8+.
- Modulo `ExchangeOnlineManagement` disponibile nel profilo PowerShell.
- Moduli Microsoft Graph per funzionalità opzionali avanzate.

## Build

```powershell
pwsh ./build/build.ps1
```

Parametri principali:

- `-Configuration Debug|Release`
- `-Clean`
- `-Publish`
- `-SelfContained`
- `-RuntimeIdentifier win-x64`
- `-Msi`
- `-ExportDirPath <path>`
- `-ImportDirPath <path>`

## Clean

```powershell
pwsh ./build/clean.ps1
```

Parametri principali:

- `-DryRun`
- `-All`
- `-SkipDotNetClean`
- `-IncludeExports`
- `-IncludeImports`
- `-ExportDirPath <path>`
- `-ImportDirPath <path>`

## Esecuzione

1. Build/publish.
2. Avvio `ExchangeAdmin.Presentation.exe`.
3. Connessione Exchange Online dalla dashboard.
4. Navigazione sezioni operative e gestione attività.
