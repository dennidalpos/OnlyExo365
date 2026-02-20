# OnlyExo365 / ExchangeAdmin

Applicazione desktop **WPF** per amministrazione operativa di **Exchange Online** con worker PowerShell separato.

## Panoramica

OnlyExo365 è pensata per operatori IT che vogliono gestire attività ricorrenti di tenant Microsoft 365/Exchange Online senza alternare continuamente portale web e script manuali.

L’app separa:
- **UI (Presentation)**: input utente, griglie, filtri, feedback visivo.
- **Worker PowerShell**: esecuzione cmdlet Exchange/Graph in runspace dedicato.
- **Infrastructure IPC**: comunicazione resiliente tra UI e worker.

Questo approccio riduce blocchi UI durante operazioni lunghe, semplifica la gestione errori e centralizza i log.

## Struttura repository

- `src/ExchangeAdmin.Presentation` — viste WPF, ViewModel, comandi, filtri lato client.
- `src/ExchangeAdmin.Application` — orchestrazione use case.
- `src/ExchangeAdmin.Infrastructure` — IPC e supervisione processo worker.
- `src/ExchangeAdmin.Worker` — cmdlet PowerShell, mapping DTO, logica backend operativa.
- `src/ExchangeAdmin.Contracts` — DTO/contratti condivisi.
- `src/ExchangeAdmin.Domain` — modelli dominio e normalizzazione errori.
- `build/` — script PowerShell di build/publish/clean.
- `installer/` — sorgenti WiX MSI.

## Prerequisiti

- Windows 10/11 o Windows Server (WPF).
- PowerShell 7+ (`pwsh` disponibile in PATH).
- .NET SDK 8+.
- Permessi amministrativi adeguati su Exchange Online / Microsoft Graph.

## Installazione moduli PowerShell

Eseguire da PowerShell 7 (consigliato amministratore):

```powershell
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force -AllowClobber
Install-Module Microsoft.Graph.Groups -Scope CurrentUser -Force -AllowClobber
```

Verifica:

```powershell
Get-Module ExchangeOnlineManagement -ListAvailable
Get-Module Microsoft.Graph* -ListAvailable
```

## Build script (`build/build.ps1`)

Esempio standard:

```powershell
pwsh ./build/build.ps1
```

### Parametri principali

- `-Configuration Debug|Release` (default: `Release`)
- `-Clean` (default: `true`)
- `-Publish` (default: `true`)
- `-SelfContained` (default: `true`)
- `-RuntimeIdentifier` (default: `win-x64`)
- `-Msi` (default: `false`)
- `-ExportDirPath <path>`
- `-ImportDirPath <path>`

### Esempi utili

Build veloce senza publish:

```powershell
pwsh ./build/build.ps1 -Publish:$false
```

Build framework-dependent:

```powershell
pwsh ./build/build.ps1 -SelfContained:$false
```

Build + MSI:

```powershell
pwsh ./build/build.ps1 -Msi
```

## Clean script (`build/clean.ps1`)

Esempio standard:

```powershell
pwsh ./build/clean.ps1
```

### Parametri principali

- `-DryRun` (simula senza cancellare)
- `-All` (pulizia estesa incluse cache/exports/imports)
- `-SkipDotNetClean`
- `-IncludeExports`
- `-IncludeImports`
- `-ExportDirPath <path>`
- `-ImportDirPath <path>`

### Comportamento aggiornato

- La pulizia `artifacts/` **preserva `exports` e `imports`** per default.
- `exports/imports` vengono rimossi automaticamente con `-All` oppure esplicitamente con `-IncludeExports` / `-IncludeImports`.
- I path custom di export/import relativi vengono risolti rispetto alla root repository.

## Logica checkbox e filtri (analisi operativa)

Questa sezione descrive il comportamento effettivo dei principali controlli UI.

### 1) Mailbox list

- **Filtro tipo mailbox (`RecipientTypeFilter`)**
  - valori supportati: `UserMailbox`, `SharedMailbox`, `RoomMailbox`, `EquipmentMailbox`, `All`.
  - il filtro viene normalizzato prima della request worker.
  - `All` invia `RecipientTypeDetails = null` (nessun filtro tipo).
- **Ricerca testo (`SearchQuery`)**
  - debounce 300 ms lato UI;
  - trim automatico;
  - reset a `null` se vuota.
- **Paginazione (`Load more`)**
  - abilitata solo con `HasMore && !IsLoading`.

### 2) Deleted mailboxes

- **Filtro UPN** (`Verifica UPN`): imposta ricerca puntuale su UPN/identità.
- **Checkbox `Includi soft-deleted`**:
  - se attiva, include mailbox soft-deleted.
  - se disattiva, esclude mailbox soft-deleted.
- **Checkbox `Includi inactive`**:
  - se attiva, include mailbox inactive.
  - se disattiva, esclude mailbox inactive.
- Cambio checkbox: trigger refresh automatico.

### 3) Distribution lists

- **Checkbox `Includi dinamiche` (`IncludeDynamic`)**
  - controlla inclusione gruppi dinamici nel listing.
  - cambio valore → refresh automatico con debounce ricerca indipendente.
- **Ricerca (`SearchQuery`)**
  - debounce 300 ms;
  - invio stringa trimmata.
- **Checkbox `Consenti mittenti esterni`**
  - legata a `RequireSenderAuthenticationEnabled` in modo inverso:
    - checked = mittenti esterni consentiti;
    - unchecked = solo mittenti autenticati.

### 4) Message trace

- **Filtro stato**: `All`, `Delivered`, `Failed`, `Pending`.
- Valori non ammessi vengono normalizzati a `All`.
- L’export usa la vista filtrata corrente.

### 5) Log viewer

- **Filtro livello minimo** (`FilterLevel`) applicato in cascata.
- **Filtro testuale** (`SearchFilter`) case-insensitive.
- Debounce e cache interna per evitare refresh UI eccessivi.

## Avvio applicazione

1. Eseguire build/publish.
2. Avviare `ExchangeAdmin.Presentation.exe` dalla cartella publish.
3. Stabilire connessione Exchange Online dalla dashboard.
4. Usare i moduli operativi (mailbox, distribution list, message trace, log).
