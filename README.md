# OnlyExo365 / ExchangeAdmin

Applicazione desktop WPF per amministrazione operativa Exchange Online con worker PowerShell separato, IPC locale e organizzazione a layer (`Contracts`, `Domain`, `Application`, `Infrastructure`, `Presentation`, `Worker`).

## Obiettivo del progetto

L'app fornisce un pannello unico per attività amministrative ricorrenti:

- Connessione a Exchange Online (e Microsoft Graph quando disponibile).
- Gestione mailbox, shared mailbox, mail flow, distribution list.
- Query su message trace con export Excel.
- Visualizzazione centralizzata dei log applicativi.

L'architettura separa interfaccia e operazioni PowerShell per migliorare isolamento degli errori, resilienza e tracciabilità.

## Architettura

- **ExchangeAdmin.Presentation (WPF)**: UI, ViewModels, filtri, validazioni input, comandi utente.
- **ExchangeAdmin.Application**: use case e service contracts verso il worker.
- **ExchangeAdmin.Infrastructure**: trasporto IPC e supervisione processo worker.
- **ExchangeAdmin.Worker**: esecuzione PowerShell (runspace), integrazione cmdlet EXO/Graph.
- **ExchangeAdmin.Contracts**: DTO e messaggi IPC condivisi.
- **ExchangeAdmin.Domain**: tipologie dominio, error taxonomy, retry/circuit-breaker.

## Flussi principali

### 1) Connessione

1. Lato UI viene richiesto il collegamento.
2. Il worker inizializza runspace PowerShell.
3. Se disponibile, viene importato `ExchangeOnlineManagement`.
4. Viene eseguito `Connect-ExchangeOnline` (opzionalmente con ambiente EXO specifico via variabile `EXCHANGEADMIN_EXO_ENV`).
5. In parallelo, viene tentata connessione Microsoft Graph per funzionalità licensing/amministrative avanzate.

### 2) Operazioni amministrative

- La UI invia una richiesta tipizzata (DTO/Message).
- Il worker esegue lo script/cmdlet e normalizza output/errori.
- Le risposte tornano via IPC con eventi progresso/log.

### 3) Export Message Trace

- La pagina Message Trace applica filtri locali sul set ricevuto.
- L'export Excel usa formato `.xlsx` OpenXML.
- La directory predefinita di export è `%LOCALAPPDATA%\OnlyExo365\exports`.

## Logica filtri e checkbox

### Distribution Lists

- **`IncludeDynamic`**: determina se includere gruppi dinamici nelle ricerche.
- La ricerca testuale usa debounce (300 ms) per ridurre chiamate ridondanti.
- Ogni cambio filtro avvia refresh sicuro con gestione eccezioni e cancellazione operazioni precedenti.
- Le checkbox impostazioni lista (es. mittenti esterni consentiti) aggiornano lo stato `HasPendingSettingsChanges` in base al diff con valori originali.

### Message Trace

- **Filtro stato** (`All`, `Delivered`, `Failed`, `Pending`) applicato localmente sulla collezione completa (`AllMessages`).
- Normalizzazione filtro con trim e confronto case-insensitive.
- Aggiornamento stato comandi dopo ogni applicazione filtro (utile per abilitazione/disabilitazione export).

### Logs

- Filtro per livello minimo (`FilterLevel`) + ricerca testuale (`SearchFilter`).
- Cache interna dei risultati filtrati invalidata ad ogni modifica rilevante.
- Refresh con debounce e soglia minima di intervallo per evitare aggiornamenti UI eccessivi.

## Prerequisiti

- **Windows** (WPF).
- **PowerShell 7+** (`pwsh`).
- **.NET SDK 8+** per build locale.
- Modulo **ExchangeOnlineManagement** installato nel profilo PowerShell.
- Per alcune funzioni, modulo **Microsoft.Graph.Authentication**.

## Build e clean

Sono presenti script PowerShell in `build/`.

### Build

```powershell
pwsh ./build/build.ps1
```

Opzioni principali:

- `-Configuration Debug|Release`
- `-Clean`
- `-Publish`
- `-SelfContained`
- `-RuntimeIdentifier win-x64`
- `-Msi`
- `-ExportDirPath <path>` per impostare la directory export artefatti

### Clean

```powershell
pwsh ./build/clean.ps1
```

Opzioni principali:

- `-DryRun` (simulazione)
- `-All` (include cache aggiuntive e anche export)
- `-IncludeExports` (pulizia esplicita export)
- `-SkipDotNetClean`

## Esecuzione applicazione

1. Build/publish del progetto.
2. Avvio `ExchangeAdmin.Presentation.exe`.
3. Connessione a Exchange Online dalla UI.
4. Navigazione delle sezioni operative (Mailbox, Mail Flow, Distribution Lists, Message Trace, Logs).

## Struttura repository

- `src/`: codice sorgente per tutti i layer.
- `build/`: script automazione build/clean.
- `installer/`: sorgenti WiX per MSI.
- `ExchangeAdmin.sln`: solution principale.
