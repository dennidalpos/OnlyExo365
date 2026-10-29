# Istruzioni del repository: OnlyExo365

<!-- secondbrain:repository:begin -->
Direttive comuni applicate: revisione **4.0** — 2026-10-29. <!-- secondbrain:source global=7585B84CED201216722B7E6646C3D2845045FEEE7E44C889E92E68A5E1975A4A workstation=2BC7FB0217FB690B8CEAB7E3DD85AAB916BA4CF3016533BC9E1D02477CE9321B -->SecondBrain: integrazione opzionale della conoscenza e delle direttive comuni. Risolvi SECONDBRAIN_HOME dallʼambiente utente della postazione; nessun percorso personale è conservato in questo repository. Se il clone privato non è disponibile, prosegui con le istruzioni tecniche locali e segnala lʼintegrazione non configurata; non creare memorie o stub sostitutivi. Dal clone SecondBrain usa scripts/prompts.ps1 -Context -RepositoryPath <checkout-corrente> per risolvere il dossier e il suo index.md. Carica indice, contesto sintetico, memoria essenziale e proiezione task pertinente; approfondimenti e storico soltanto quando servono. Memorie private, tracker interni, backup e output operativi risiedono in SecondBrain. Conserva qui codice, guide e asset pubblicabili. Le direttive non autorizzano commit, push, cancellazioni di artefatti o task applicativi impliciti.
<!-- secondbrain:repository:end -->

Revisione specifica del progetto 2.5 — 2026-10-25. Vincoli specifici del progetto; i contenuti comuni sono gestiti separatamente.

## Contesto

- **Scopo**: Applicazione desktop Windows x64 per l'amministrazione, reporting e gestione di tenant Microsoft Exchange Online e Microsoft Graph: Shell WPF di presentazione, Worker isolato per l'esecuzione di script PowerShell e libreria Contracts per il protocollo condiviso.
- **Modalità**: Greenfield (semplificazione aggressiva per architettura minimale; eliminazione shim temporanei e dead code)
- **Target Applicativo**: Desktop Windows x64 (WPF / .NET x64).
- **Stack & Toolchain**: C# / .NET, WPF, PowerShell SDK / Exchange Online Management, Microsoft Graph API, xUnit, `dotnet` e PowerShell (`pwsh`).

## Vincoli specifici

  * Autenticazione Graph/Exchange sicura: nessuna credenziale, certificato o token tenant salvato in chiaro nel repository.
  * **Segreti**: `.env.example` contiene esclusivamente placeholder fittizi.
  * Validazione rigorosa degli input prima di inoltrare comandi al processo Worker.
  * Separazione netta dei processi: la Shell gestisce solo la UI; il Worker gestisce l'esecuzione PowerShell/Graph.
  * Comunicazione IPC tipizzata e validata tramite contratti condivisi (`OnlyExo365.Contracts`).
  * Timeout stringenti e propagazione di `CancellationToken` per tutte le operazioni di rete o tenant.

## Comandi e verifica

Comandi verificati staticamente contro i file del progetto; questa revisione delle direttive non certifica l'esecuzione di build o test applicativi. Consultare anche i gate della CI, quando presenti.

| Attività | Comando | Ambiente | Evidenza e limiti |
|---|---|---|---|
| **Bootstrap** | `pwsh ./scripts/bootstrap.ps1 -RuntimeIdentifier win-x64` | pwsh / root | Configurazione ambiente e runtime |
| **Build Debug** | `pwsh ./scripts/build.ps1 -Configuration Debug -RuntimeIdentifier win-x64` | pwsh / root | Compilazione soluzione in Debug |
| **Start / Run** | `pwsh ./scripts/start.ps1 -Configuration Debug -RuntimeIdentifier win-x64 -NoBuild` | pwsh / root | Avvio locale Shell per collaudo |
| **Targeted checks** | `dotnet test --filter "FullyQualifiedName~Target" --nologo` | pwsh / root | Test rapidi mirati all'ambito modificato |
| **Full checks** | `dotnet test --nologo && pwsh ./scripts/build.ps1 -Configuration Release -RuntimeIdentifier win-x64` | pwsh / root | **Gate obbligatorio pre-commit** (test + build Release) |

## Riferimenti e collaborazione

- Consultare README e documentazione effettivamente presenti; versioni dai manifest/lockfile, non da ipotesi sullo stack.
- 02-areas/products/OnlyExo365/memory.md e 02-areas/products/OnlyExo365/context.md (nel dossier SecondBrain) conservano decisioni tecniche durevoli, gotcha e architettura; aggiornarli a fine task.
- Tracker autorevole nel dossier SecondBrain: `02-areas/products/OnlyExo365/tracker.json`, dichiarato in repository.json. Conservare schema JSON v1 e storico; build e test restano indipendenti da SecondBrain.
- Eseguire i controlli previsti per la modifica prima di un commit richiesto. Nessun commit o push automatico.
- Preservare dati persistenti e configurazioni esterne; non estendere la manutenzione a operazioni distruttive non autorizzate.
