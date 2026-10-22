# Istruzioni del repository: OnlyExo365

<!-- secondbrain:repository:begin -->
Direttive comuni applicate: revisione **3.0** — 2026-10-19.
<!-- secondbrain:source global=DC59C089490D9EE3E7E7F389FCC087D199EB12C16ADD2D20EA78076D9BB3BBDB workstation=3090606D3A18890581CDAD79BE8A65348612B40AAF40D5EAEB4E62AE0215D430 -->
Dossier SecondBrain di questo repository: 02-areas/products/OnlyExo365, relativo alla radice di SecondBrain indicata nelle istruzioni globali.
- Leggi a inizio task: questo AGENTS.md, 02-areas/products/OnlyExo365/context.md, 02-areas/products/OnlyExo365/memory.md e PROJECT_STATUS.json (o 02-areas/products/OnlyExo365/tracker.md).
- Nei task di scrittura aggiorna solo novità pertinenti: decisioni e gotcha in 02-areas/products/OnlyExo365/memory.md; architettura e stato in 02-areas/products/OnlyExo365/context.md; attività nel tracker preservandone schema e storico.
- Nei task di scrittura pulisci temporanei del task e residui confermati, preservando dati e lavoro unico; rivedi diff e git status --ignored. Esegui secondbrain.ps1 -Sync dopo modifiche a dossier o direttive, oppure su richiesta. Nei task di sola lettura segnala i residui senza rimuoverli.
<!-- secondbrain:repository:end -->

Revisione specifica del progetto 2.4 — 2026-10-18. Vincoli specifici del progetto; i contenuti comuni sono gestiti separatamente.

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
- PROJECT_STATUS.json è il tracker esistente: contiene lavoro aperto, senza storico dei task conclusi; preservarne lo schema.
- Eseguire i controlli previsti per la modifica prima di un commit richiesto. Nessun commit o push automatico.
- Preservare dati persistenti e configurazioni esterne; non estendere la manutenzione a operazioni distruttive non autorizzate.
