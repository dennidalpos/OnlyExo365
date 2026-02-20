# OnlyExo365 / ExchangeAdmin

Applicazione desktop **WPF** per amministrazione operativa di **Exchange Online**, progettata con architettura a layer e worker PowerShell separato per mantenere la UI reattiva durante operazioni potenzialmente lunghe.

## Obiettivo del progetto

OnlyExo365 centralizza in una singola interfaccia le attività amministrative più frequenti in ambienti Microsoft 365 / Exchange Online, riducendo il passaggio continuo tra console PowerShell, portali web e script locali.

Funzioni principali:

- connessione e gestione sessione verso Exchange Online;
- analisi mailbox (utente, shared, cancellate);
- gestione distribution list statiche e dinamiche;
- consultazione e filtraggio message trace;
- esportazione dati operativi in formato `.xlsx`;
- visibilità centralizzata dei log applicativi.

## Architettura del repository

- `src/ExchangeAdmin.Presentation`
  UI WPF, ViewModel, validazione input, binding, comandi utente, filtri locali e stato interfaccia.
- `src/ExchangeAdmin.Application`
  orchestrazione dei casi d’uso applicativi e coordinamento servizi.
- `src/ExchangeAdmin.Infrastructure`
  comunicazione IPC locale tra UI e worker, gestione ciclo vita worker e resilienza di processo.
- `src/ExchangeAdmin.Worker`
  esecuzione PowerShell in runspace dedicato, import moduli, invocazione cmdlet Exchange Online / Graph.
- `src/ExchangeAdmin.Contracts`
  DTO, contratti e messaggi condivisi tra componenti.
- `src/ExchangeAdmin.Domain`
  tipologie di dominio, errori normalizzati, regole comuni.
- `build/`
  script di build e pulizia artefatti.
- `installer/`
  sorgenti WiX per packaging MSI.

## Flusso operativo

1. La UI invia una richiesta tipizzata al layer Application.
2. Application delega la richiesta a Infrastructure.
3. Infrastructure inoltra la richiesta al worker tramite IPC locale.
4. Il worker esegue i cmdlet in PowerShell runspace isolato.
5. Risultati, errori e stati di avanzamento vengono normalizzati.
6. La UI aggiorna stato, griglie, dettagli e log operativi.

Questa separazione riduce i freeze della UI, isola errori runtime PowerShell e rende più prevedibile il comportamento dell’app in caso di fault.

## Prerequisiti

- Windows (richiesto per WPF).
- PowerShell 7+ (`pwsh`).
- .NET SDK 8+.
- Accesso amministrativo a Exchange Online.
- Moduli PowerShell necessari installati nel profilo utente o all users.

## Installazione manuale moduli PowerShell (obbligatoria)

Apri **PowerShell 7 come amministratore** ed esegui i comandi seguenti.

### 1) Verifica e prepara PSGallery

```powershell
Get-PSRepository -Name PSGallery
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
```

### 2) Installa modulo Exchange Online

```powershell
Install-Module ExchangeOnlineManagement -Scope AllUsers -Force -AllowClobber
```

Se preferisci installare solo per l’utente corrente:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
```

### 3) Installa moduli Microsoft Graph (funzionalità opzionali avanzate)

Installazione completa:

```powershell
Install-Module Microsoft.Graph -Scope AllUsers -Force -AllowClobber
```

Installazione ridotta (moduli comunemente utili):

```powershell
Install-Module Microsoft.Graph.Authentication -Scope AllUsers -Force -AllowClobber
Install-Module Microsoft.Graph.Users -Scope AllUsers -Force -AllowClobber
Install-Module Microsoft.Graph.Groups -Scope AllUsers -Force -AllowClobber
```

### 4) Verifica finale installazione

```powershell
Get-Module ExchangeOnlineManagement -ListAvailable
Get-Module Microsoft.Graph* -ListAvailable
```

## Build

Esegui dalla root del repository:

```powershell
pwsh ./build/build.ps1
```

Parametri principali supportati dallo script:

- `-Configuration Debug|Release`
- `-Clean`
- `-Publish`
- `-SelfContained`
- `-RuntimeIdentifier win-x64`
- `-Msi`
- `-ExportDirPath <path>`
- `-ImportDirPath <path>`

Esempio build release con publish:

```powershell
pwsh ./build/build.ps1 -Configuration Release -Publish
```

## Clean

Pulizia artefatti:

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

## Esecuzione applicazione

1. Completa i prerequisiti e installa i moduli PowerShell.
2. Esegui build/publish.
3. Avvia `ExchangeAdmin.Presentation.exe`.
4. Effettua la connessione a Exchange Online dalla dashboard.
5. Accedi alle aree operative (mailbox, liste, message trace, log).

## Comportamenti funzionali rilevanti

### Message Trace ed export

- filtro stato supportato: `All`, `Delivered`, `Failed`, `Pending`;
- valori non validi vengono normalizzati a `All`;
- export in `.xlsx` sui risultati filtrati correnti;
- directory di export:
  - priorità: variabile ambiente `EXCHANGEADMIN_EXPORT_DIR`;
  - fallback: `%LOCALAPPDATA%\OnlyExo365\exports`.

### Distribution Lists

- supporto gruppi statici e dinamici (`IncludeDynamic`);
- rilevamento modifiche pendenti sulle impostazioni;
- gestione coerente dei mittenti ammessi/bloccati con normalizzazione input.

### Mailbox

- filtro recipient type (`UserMailbox` / `SharedMailbox`);
- ricerca testuale normalizzata (trim/null);
- caricamento progressivo con `LoadMore` su base `HasMore` + `Skip`.

### Logging

- filtro per livello minimo;
- filtro testuale case-insensitive;
- aggiornamento lista con debounce/throttling per ridurre refresh UI eccessivi.
