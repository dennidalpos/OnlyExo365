# ExchangeAdmin

Applicazione desktop WPF per amministrare Exchange Online tramite un worker PowerShell separato e comunicazione IPC su Named Pipes.

## Requisiti

**Runtime utente**
- Windows 10/11 (WPF è Windows‑only).
- PowerShell 7+ disponibile in PATH (`pwsh`).
- Modulo PowerShell `ExchangeOnlineManagement` installato.
- .NET 8 Runtime (solo per build framework‑dependent).

**Ambiente di sviluppo**
- .NET 8 SDK.
- PowerShell 7+ (per gli script di build).

## Setup

1. Installa PowerShell 7 e verifica:
   ```powershell
   pwsh --version
   ```
2. Installa il modulo Exchange Online:
   ```powershell
   pwsh -Command "Install-Module ExchangeOnlineManagement -Scope CurrentUser"
   ```
3. Verifica il modulo:
   ```powershell
   pwsh -Command "Get-Module -ListAvailable ExchangeOnlineManagement"
   ```

## Comandi principali

### Sviluppo (build + publish + avvio)
```powershell
pwsh ./build/quick-test.ps1
```
Opzioni: `-SkipClean`, `-SkipBuild`, `-LaunchOnly`.

### Build produzione
```powershell
pwsh ./build/build.ps1 -Clean -Publish
```

### Build self‑contained
```powershell
pwsh ./build/build.ps1 -Clean -Publish -SelfContained
```

### MSI (richiede WiX Toolset v3.14)
```powershell
pwsh ./build/build.ps1 -Publish -Msi
```

### Test
Non ci sono progetti di test automatizzati nel repository.

### Lint
Non è presente un comando di lint dedicato.

## Configurazione (variabili d’ambiente)

- `EXCHANGEADMIN_EXO_ENV`: imposta l’ambiente di Exchange Online usato dal worker (passato al comando di connessione).
- `EXCHANGEADMIN_DISABLE_EXO`: se impostata a `1|true|yes|on`, disabilita la connessione Exchange nell’interfaccia.

## Struttura cartelle

```
build/                      Script di build e quick-test
installer/                  Definizione MSI (WiX)
src/
  ExchangeAdmin.Application
  ExchangeAdmin.Contracts
  ExchangeAdmin.Domain
  ExchangeAdmin.Infrastructure
  ExchangeAdmin.Presentation
  ExchangeAdmin.Worker
ExchangeAdmin.sln
README.md
TESTING.md
```

## Esecuzione in locale

1. Esegui la build/publish con `build/quick-test.ps1` oppure `build/build.ps1`.
2. Avvia `artifacts/publish/ExchangeAdmin.Presentation.exe`.
3. Dal client:
   - avvia il worker,
   - connettiti a Exchange Online (verrà aperta una finestra browser per l’autenticazione).

## Esecuzione in produzione

1. Genera il publish (`build/build.ps1 -Publish` o `-SelfContained`).
2. Distribuisci la cartella `artifacts/publish`.
3. Avvia `ExchangeAdmin.Presentation.exe` nella cartella di publish: il worker viene lanciato automaticamente.

## Troubleshooting essenziale

- **PowerShell non trovato**: assicurati che `pwsh` sia in PATH.
- **Modulo ExchangeOnlineManagement assente**: installalo con `Install-Module ExchangeOnlineManagement`.
- **Worker non parte**: verifica che `ExchangeAdmin.Worker.exe` sia accanto a `ExchangeAdmin.Presentation.exe` nella cartella di publish.
- **Autenticazione bloccata**: la finestra del worker è minimizzata ma deve poter aprire il browser; non disabilitare la finestra.
