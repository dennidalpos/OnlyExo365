# OnlyExo365 Documentation Index

Guide per amministratori Exchange Online, help desk e manutentori. Per iniziare: [README del prodotto](../README.md), [guida operatore](operations/operator-guide.md) e [percorsi visivi](operations/visual-guide.md). Le funzionalità dipendono dalle capacità e dai permessi della sessione.

## Operatività

| Documento | Contenuto | Fonti principali |
| --- | --- | --- |
| [Guida operatore](operations/operator-guide.md) | Installazione, prima connessione, modifiche, Message Trace, flusso posta e assistenza | `Shell/ViewModels`, `ExoMessageTraceCommands.cs` |
| [Percorsi visivi](operations/visual-guide.md) | Otto flussi Mermaid: connessione, Dashboard, cassette, trace, sicurezza, conformità, flusso posta e diagnostica | `Shell/Views`, `Shell/ViewModels` |
| [Moduli e autenticazione](operations/authentication.md) | Versioni esatte, modalità e limiti Exchange/Graph/Compliance | `Contracts/Data/PowerShellModuleBootstrapPolicy.json`, `Contracts/Data/CloudEnvironments.json` |
| [Diagnostica](operations/diagnostics.md) | Console Worker, log persistenti ed esportazione | `WorkerConsoleController.cs`, `PersistentLogStore.cs` |
| [Asset pubblici](assets/README.md) | Schermata reale, versione, scenario e riproduzione | `assets/dashboard-disconnected.png` |
| [Copy per locandine](operations/locandine-copy.md) | Testi informativi e limiti da mantenere nel materiale pubblico | README, guide e asset pubblici |

## Configurazione e dati

| Documento | Contenuto | Fonti principali |
| --- | --- | --- |
| [Configurazione](configuration/settings.md) | Default, file dell'applicazione, policy macchina e override di ambiente | `ExchangeConfigurationLoader.cs`, `ExchangeOnlineConfiguration.cs` |
| [Archiviazione locale](configuration/storage.md) | Percorsi, log, riferimenti DPAPI, report atomici, preferenze e disinstallazione | `PersistentLogWriter.cs`, `AtomicFileWriter.cs`, `installer/OnlyExo365.iss` |
| [Catalogo licenze](licensing/catalog.md) | Catalogo incorporato, cache locale, aggiornamento e parsing CSV | `LicenseCatalogDownloader.cs`, `Microsoft365SkuCatalog.json` |

## Architettura e sviluppo

| Documento | Contenuto | Fonti principali |
| --- | --- | --- |
| [Architettura](architecture/overview.md) | Shell WPF, Worker isolato, Contracts e sequenza delle richieste | `src/`, `Shell/Bootstrap`, `Worker/Operations` |
| [IPC](architecture/ipc.md) | Named pipe, autenticazione sessione, framing, limiti e cancellazione | `Contracts/IpcLineReader.cs`, `Contracts/IpcConstants.cs`, `Worker/Ipc/IpcServer.cs` |
| [Workflow di sviluppo](development/workflow.md) | Toolchain vincolata, restore, build, test e gate | `global.json`, `scripts/`, `build/` |
| [Qualifica degli ambienti](development/qualification.md) | Verifiche su VM, tenant dedicato e UI interattiva; limiti delle prove offline | `run-smoke-tests.ps1`, `run-tenant-validation.ps1` |
| [Packaging](releases/packaging.md) | Installer, pacchetti e bundle di rilascio | `scripts/pack.ps1`, `installer/OnlyExo365.iss` |
| [Firma digitale](releases/signing.md) | Sequenza di firma e verifica Authenticode | `build/sign-artifacts.ps1`, `build/verify-signatures.ps1` |

I nomi di sorgente nelle tabelle sono riferimenti sintetici; i collegamenti alle guide portano ai percorsi completi.

## Audit storici

Le note seguenti descrivono le revisioni indicate, non la qualifica della versione corrente. Queste pagine conservano sintesi tecniche o rimandi: i rapporti integrali e le prove operative sono nel dossier del progetto e non in questa raccolta pubblica.

- [Affidabilità critica — 10/10/2026](audits/2026-10-10-critical-audit.md).
- [UI/UX — 10/10/2026](audits/2026-10-10-ui-ux-audit.md).
- [PowerShell — 10/10/2026](audits/2026-10-10-powershell-audit.md).
- [Fixture conservate — 10/10/2026](audits/2026-10-10-retained-fixtures.md).
- [Revisione documentazione — 10/10/2026](audits/2026-10-10-documentation-review.md).
- [Rimedi locali — 11/10/2026](audits/2026-10-11-local-remediation.md).
- [Mail flow P1 — 26/10/2026](audits/2026-10-26-mailflow-p1.md).
- [Dipendenze NuGet — 26/10/2026](audits/2026-10-26-nuget-remediation.md).
- [Rimedi P2 — 27/10/2026](audits/2026-10-27-p2-remediation.md).
- [Audit applicazione — 27/10/2026](audits/2026-10-27-application-audit.md).

Per i requisiti delle prove correnti consultare [qualifica](development/qualification.md); per il comportamento usare le guide e le sorgenti della revisione in esame.
