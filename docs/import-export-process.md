# Import / Export Processes

## Export Message Trace (Excel)

- Entry point: `MessageTraceViewModel.ExportExcelCommand`.
- UI action: bottone **Export Excel** nella pagina Message Trace.
- Formato output: `.xlsx` (OpenXML), non CSV.
- Contenuto esportato: righe attualmente visualizzate (`Messages`), quindi rispettando filtro status client-side.
- Colonne:
  1. Received
  2. Sender
  3. Recipient
  4. Subject
  5. Status
  6. Size
  7. MessageId
  8. MessageTraceId
- Formattazione:
  - header con stile dedicato (bold + fill).
  - valori serializzati in testo per massima compatibilità e audit readability.

## Import modules (runtime prerequisites)

- Check prerequisiti: `ToolsViewModel.CheckPrerequisitesCommand`.
- Install/aggiornamento moduli: `ToolsViewModel.InstallExchangeModuleCommand` e `InstallGraphModuleCommand`.
- Esecuzione worker:
  - `CheckPrerequisites` verifica PowerShell + moduli Exchange/Graph.
  - `InstallModule` esegue installazione in scope utente.
- Log operativi:
  - prefisso `[Prerequisites]` per check.
  - prefisso `[ModuleInstall]` per install/update.

## Clean/build integration

- `build/build.ps1` crea anche `artifacts/exports` come cartella standard per output export.
- `build/clean.ps1` supporta `-IncludeExports` per pulire esplicitamente export generati.

## Note operative

- Nessun changelog automatico: la documentazione tecnica viene mantenuta nei file `docs/`.
- In ambienti senza SDK .NET locale, validare script/flow con controlli statici e CI.
