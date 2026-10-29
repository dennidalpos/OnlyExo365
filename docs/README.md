# OnlyExo365 Documentation Index

Guida modulare alla documentazione del progetto, organizzata per ambiti funzionali e operativi. I rapporti di audit registrano evidenze storiche relative alle revisioni di riferimento; per lo stato attivo e i requisiti aperti consultare 02-areas/products/OnlyExo365/tracker.json (nel dossier SecondBrain) (SecondBrain: `02-areas/products/OnlyExo365/tracker.json`).

---

## 1. Architettura & Protocollo

| Documento | Descrizione | Sorgenti correlate |
| --- | --- | --- |
| [Panoramica architetturale](architecture/overview.md) | Responsabilità dei componenti (.NET 10 x64, Shell WPF, Worker isolato, Contracts, Tests), composizione eager dei ViewModel delle 18 pagine e vincoli di modularità. | `src/OnlyExo365.Shell/Bootstrap`, `build/assert-architecture-constraints.ps1` |
| [Protocollo IPC](architecture/ipc.md) | Comunicazione Windows Named Pipes duplex, isolamento processo worker, session token, vincolo di dimensione messaggi (10 MB via `IpcLineReader` con chunk 256 KiB), timeout e gestione eventi. | `src/OnlyExo365.Contracts/IpcConstants.cs`, `src/OnlyExo365.Contracts/Ipc/IpcLineReader.cs`, `src/OnlyExo365.Worker/Ipc/IpcServer.cs`, `src/OnlyExo365.Shell/Ipc/WorkerSupervisor.cs` |

## 2. Configurazione & Archiviazione locale

| Documento | Descrizione | Sorgenti correlate |
| --- | --- | --- |
| [Configurazione ed ambiente](configuration/settings.md) | Gerarchia di configurazione a 3 livelli (defaults, `appsettings.json`, machine policy, variabili d'ambiente `ONLYEXO365_*`), permessi Graph e inibizione connessione (`ONLYEXO365_DISABLE_EXO`). | `src/OnlyExo365.Shell/Services/ExchangeConfigurationLoader.cs`, `src/OnlyExo365.Contracts/ExchangeOnlineConfiguration.cs` |
| [Archiviazione locale](configuration/storage.md) | Mappa dei percorsi su filesystem, log persistenti (`%LocalAppData%\OnlyExo365\logs`, retention 14 giorni), segreti DPAPI, report XLSX con scrittura atomica (`AtomicFileWriter`), preferenze utente in Roaming AppData e contratti dati per clean e disinstallazione. | `src/OnlyExo365.Contracts/Diagnostics/PersistentLogWriter.cs`, `src/OnlyExo365.Contracts/AtomicFileWriter.cs`, `src/OnlyExo365.Shell/Services/UserPreferencesService.cs`, `installer/OnlyExo365.iss` |

## 3. Licensing & Catalogo SKU

| Documento | Descrizione | Sorgenti correlate |
| --- | --- | --- |
| [Catalogo licenze Microsoft 365](licensing/catalog.md) | Baseline incorporata, parsing CSV standard con `TextFieldParser`, cache locale, frequenze di refresh (Daily/Monthly) e script di rigenerazione. | `src/OnlyExo365.Shell/Services/LicenseCatalogDownloader.cs`, `src/OnlyExo365.Worker/Data/Microsoft365SkuCatalog.json`, `scripts/agents/refresh-microsoft365-sku-catalog.ps1` |

## 4. Operatività & Amministrazione

| Documento | Descrizione | Sorgenti correlate |
| --- | --- | --- |
| [Guida operatore](operations/operator-guide.md) | Prerequisiti minimi, procedura di primo avvio e connessione, limiti e paginazione di Message Trace (intervallo 10 giorni / 90 giorni, batch da 5.000 con lookahead page+1). | `src/OnlyExo365.Worker/PowerShell/ExoMessageTraceCommands.cs`, `src/OnlyExo365.Contracts/Dtos/MessageTraceDto.cs` |
| [Guida visiva percorsi utente](operations/visual-guide.md) | Diagrammi di flusso Mermaid per tutte le funzionalità chiave: connessione, dashboard, cassette postali, message trace, sicurezza posta, conformità, flusso posta e strumenti. | `src/OnlyExo365.Shell/Views`, `src/OnlyExo365.Shell/ViewModels` |
| [Materiale di comunicazione e copy](operations/locandine-copy.md) | Copy per locandine, roll-up fieristici (85x200), bacheca IT A3/A4, flyer fronte-retro e card social con riferimenti alle schermate PNG live. | `docs/assets/screenshots`, `docs/operations/locandine-copy.md` |
| [Moduli e autenticazione](operations/authentication.md) | Politica di bootstrap per versioni esatte dei moduli approvati (EXO 3.10.1, Graph 2.41.2), modalità di connessione (Interactive, DeviceCode, AppCertificate, ManagedIdentity) e vincoli su sessioni Compliance/eDiscovery (`Connect-IPPSSession`). | `src/OnlyExo365.Contracts/Data/PowerShellModuleBootstrapPolicy.json`, `src/OnlyExo365.Worker/PowerShell/ComplianceCommandBuilder.cs` |
| [Diagnostica](operations/diagnostics.md) | Controllo e monitoraggio del Worker tramite la pagina Tools, attivazione console runtime, ispezione/esportazione log persistenti e corretta gestione della cultura di conversione UI (`ConverterCulture=en-US`). | `src/OnlyExo365.Worker/WorkerConsoleController.cs`, `src/OnlyExo365.Shell/Views/MainWindow.xaml` |

## 5. Sviluppo & Ciclo di vita

| Documento | Descrizione | Sorgenti correlate |
| --- | --- | --- |
| [Workflow di sviluppo](development/workflow.md) | Baseline vincolata (.NET SDK 10.0.403 pinned con rollForward disabilitato, Inno Setup 6.7.1), comandi locali di restore, compilazione, esecuzione gate e disciplina di pulizia controllata. | `global.json`, `scripts/bootstrap.ps1`, `scripts/gate.ps1`, `scripts/clean.ps1`, `build/assert-architecture-constraints.ps1` |
| [Qualifica degli ambienti](development/qualification.md) | Procedure di collaudo avanzato: smoke test installer su VM pulita/disposable, validazione tenant su tenant dedicato e criteri di verifica interattiva UI/accessibilità (UX-08). | `build/run-smoke-tests.ps1`, `build/run-tenant-validation.ps1` |

## 6. Rilascio & Distribuzione

| Documento | Descrizione | Sorgenti correlate |
| --- | --- | --- |
| [Packaging e asset di rilascio](releases/packaging.md) | Compilazione dell'installer con Inno Setup, pacchetti framework-dependent e self-contained, verifica della riproducibilità e generazione bundle di rilascio. | `scripts/pack.ps1`, `build/verify-package-reproducibility.ps1`, `scripts/agents/publish.ps1`, `installer/OnlyExo365.iss` |
| [Firma digitale](releases/signing.md) | Integrazione con Windows SDK SignTool, sequenza di firma obbligatoria (payload publish -> compilazione installer -> firma installer -> verifica firme). | `build/sign-artifacts.ps1`, `build/create-setup-exe.ps1`, `build/verify-signatures.ps1`, `build/validate-artifact-signing.ps1` |

## 7. Rapporti di audit e verifiche storiche (`docs/audits/`)

> [!NOTE]
> I file di audit descrivono le evidenze e le verifiche condotte sulle revisioni indicate. Non qualificano modifiche successive.

- [Rimedio affidabilità critica (2026-10-10)](audits/2026-10-10-critical-audit.md): risoluzione AUD-01..AUD-06 (gestione dati disinstallazione, binding parametri credenziali, envelope IPC e retry backoff, paginazione trace, policy unificata moduli).
- [Rimedio UI/UX e accessibilità (2026-10-10)](audits/2026-10-10-ui-ux-audit.md): risoluzione UX-01..UX-07 (debounce, dialog modali con owner, contrasto e palette, dirty tracking, virtualizzazione, correzione crash startup MainWindow).
- [Rimedio script PowerShell e automazione (2026-10-10)](audits/2026-10-10-powershell-audit.md): risoluzione PS-01..PS-11 (ripristino backup atomico, sicurezza clean e delimitazione perimetri, sequenza firma, riproducibilità binari).
- [Registro fixture trattenute: CLEAN-01 (2026-10-10)](audits/2026-10-10-retained-fixtures.md): tracciamento delle fixture di audit non rimosse per vincoli di execution policy harness.
- [Revisione documentazione e codice (2026-10-10)](audits/2026-10-10-documentation-review.md): riconciliazione su 13 documenti, rimozione vecchi ADR, identificazione difetto limite byte IPC-01.
- [Rimedio locale e chiusura task (2026-10-11)](audits/2026-10-11-local-remediation.md): chiusura formale di IPC-01 (`IpcLineReader`), IMP-01 (runspace coverage comandi generati), IMP-02 (`AtomicFileWriter`), IMP-03 (`TextFieldParser`), IMP-04 (misurazione bootstrap) e risoluzione decisione LOC-01 (conservazione diagnostica tecnica in italiano).
- [Rimedio P1 Mail flow (2026-10-26)](audits/2026-10-26-mailflow-p1.md): azzeramenti delle regole, sintassi connettori/domini e preservazione della categoria connettori; 843 test e build Release verificati offline.
- [Rimedio P2 (2026-10-27)](audits/2026-10-27-p2-remediation.md): catalogo automatico, cache per sessione, letture fallite, routing cloud, matcher a tre stati e deadline invio IPC; evidenze e limiti del lotto.
- [Audit applicazione e dipendenze (2026-10-27)](audits/2026-10-27-application-audit.md): correzioni Compliance, recovery, IPC, cancellazione e release; aggiornamenti qualificati e residui di firma/modulo EXO separati.
