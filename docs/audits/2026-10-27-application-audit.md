# Audit applicazione e aggiornamenti — 2026-10-27

Base: `a1ca489c9e8817b8fc416b5842c0d0a88dd8de1c`, checkout inizialmente pulito. Esecuzione locale Windows x64, senza tenant, installazione/disinstallazione, modifiche alle impostazioni personali, commit, push o pubblicazione.

## Risultato e perimetro verificato

Correzioni locali e aggiornamenti qualificati con test offline. L'audit segue le superfici critiche Shell/Worker/Contracts e i flussi di autenticazione, richieste, generazione comandi, letture, cache, diagnostica ed esportazione. I test con cmdlet simulati verificano il codice applicativo, senza certificare il comportamento dei servizi Microsoft 365.

| Area | Evidenze e limiti |
| --- | --- |
| Correttezza | Regressioni per selezione certificati, stato dopo recovery, argomenti di release; suite esistente per mapping, letture fallite, aggiornamenti Mail flow, cache e routing cloud. Non tutte le combinazioni dei cmdlet sono eseguite su servizi reali. |
| Affidabilità | Cancellazione di importazione, recovery, HTTP bloccato e processi ausiliari; capacità IPC e invii bloccati; test di coda, deadline, retry e messaggi sovradimensionati. |
| Sicurezza | Parametri release separati dal codice PowerShell, trasporto autenticato/current-user-only, password parametrizzate/DPAPI, memoria dei buffer e diagnostica; scansione Gitleaks e audit NuGet transitivo. Il residuo di firma è separato sotto. |
| Prestazioni | Carichi sintetici riproducibili per mapping, raccolte, log e XLSX; Message Trace mantiene pagina più lookahead e limita le chiamate negli scenari da 10.000 elementi e pagina profonda. Nessun profiling UI reale. |
| Dipendenze | Manifest, quattro lockfile, artefatti e probe del runtime incorporato; confronto con fonti ufficiali esatte. NuGet non analizza automaticamente tutte le DLL distribuite dentro i moduli PowerShell. |

## Problemi corretti

| Priorità | Problema | Correzione / verifica |
| --- | --- | --- |
| P1 | Tag di release interpolati direttamente nel sorgente `run`, anche nei passaggi successivi alla validazione. | Variabili d'ambiente in tutti i consumatori, validazione Git e rifiuto di newline/opzioni. Tag con apici, `$()` e `;` restano dati. Nessuna esecuzione remota del workflow. |
| P2 | `REV-20261027-01`: Compliance subject-only usa un parametro assente da EXO 3.10.1. | Risoluzione esatta in `CurrentUser/My`, un solo certificato attualmente valido con chiave privata, invio del thumbprint. Precedenza del thumbprint esplicito. Stesse regole nello script tenant; fixture senza store personale. |
| P2 | Recovery conserva Compliance connessa e scope Graph del runspace scartato; può precedere il lock di esecuzione. | Azzeramento di tutti gli stati prima della ricreazione; recovery protetta dal lock e dalla cancellazione. Test riprodotto in errore prima della correzione (`recovery-before.trx`). |
| P2 | Inizializzazione/importazione/ricreazione non ricevono il token del chiamante. | Token propagato, budget di 30 secondi, arresto della pipeline e distinzione fra timeout interno e cancellazione volontaria. `InitializeAsync` accetta un token opzionale, mantenendo compatibili i chiamanti sorgente. |
| P2 | Risposte/eventi Worker possono attendere senza limite lock, scrittura o flush; dispatch concorrente aggira la sola capacità della coda. | Budget di invio di 5 secondi; abbandono del trasporto su cancellazione. Massimo 9 richieste ammesse complessivamente, incluse quelle concorrenti; sovraccarico restituisce `ServiceUnavailable`. |
| P2 | Il client può riutilizzare una connessione dopo cancellazione durante un frame parziale. | Disconnessione quando write/flush viene interrotto; attesa del lock o della risposta mantiene la connessione. Chiusura pipe prima dell'attesa dei loop e dello smaltimento del writer. Nessun retry automatico delle scritture tenant. |
| P2 | Annullare il bootstrap interattivo dispone il `Process` ma lascia vivo l'host di autenticazione; il probe execution-policy non ha timeout locale. | Terminazione del solo albero di processo posseduto, con attesa finale limitata. Budget di 5 minuti per sign-in e 5 secondi per il probe; test con figli innocui e nascosti. |
| P2 | Download HTTP nel fallback di bootstrap moduli non riceve la cancellazione durante `GetAsync`/lettura contenuto. | Token esplicito, timeout di 60 secondi e disposal della risposta. Regressione con TCP loopback che non invia risposta, senza installare moduli personali. |
| P3 | `ProtectedSecretStore.Create` azzera il buffer solo sul percorso riuscito. | Azzeramento in `finally`, anche dopo errore DPAPI/I/O, e rimozione best effort del file parziale. Roundtrip/trasporto coperti dalla suite; l'azzeramento sul ramo eccezionale è verificato staticamente. |

Contratti IPC, DTO e nomi delle opzioni di configurazione restano invariati. Timeout/cancellazione non annullano modifiche già inviate al tenant.

## Versioni rilevate e decisioni

| Componente | Prima | Dopo / decisione |
| --- | --- | --- |
| SDK .NET | Pin 10.0.401; sistema 10.0.403 | Pin 10.0.403, `rollForward: disable` conservato. |
| Runtime .NET / Desktop | Installato 10.0.14 | Conservato; probe incorporato eseguito su .NET 10.0.14. Nessun runtime installato dal task. |
| PowerShell SDK / esterno | 7.6.6 / 7.6.9 | SDK 7.6.9; esterno invariato. Assembly SMA caricata dal probe identica per SHA-256 a quella dell'artefatto Worker Windows. |
| Xml / Pkcs | 10.0.13 | 10.0.14 tramite PowerShell SDK; override Xml rimosso. |
| ProtectedData | Riferimento Contracts 10.0.12 | 10.0.14. |
| ExchangeOnlineManagement | 3.10.1 | Conservato; import reale offline e parametri Compliance verificati. |
| Graph Authentication, Users, Users.Actions, Identity.DirectoryManagement | 2.41.1 | Policy esatta 2.41.2; quattro moduli scaricati e importati solo nello scratch. L'installazione personale 2.41.1 resta invariata e richiede il normale bootstrap prima dell'uso. |
| Open XML | 3.5.1 | Conservato. |
| Microsoft.NET.Test.Sdk / coverlet.collector | 18.4.0 / 10.0.0 | 18.11.0 / 10.1.0. |
| xUnit assert/core / StaFact / adapter | 3.2.2 / 3.0.13 / 3.1.5 | 4.0.2 / 4.0.24 / 4.0.1; `core.mtp-off` mantiene VSTest. Parallelismo disabilitato con il nuovo attributo xUnit. |
| Gitleaks | 8.30.0 | 8.30.1, checksum SHA-256 ufficiale verificato; nessun cambiamento alle esclusioni. |
| PSScriptAnalyzer | 1.25.0 | Conservato. |

L'inventario JSON conserva versioni richieste, risoluzioni per framework/RID, versioni prodotto degli artefatti e assembly caricati dal probe. Una versione assembly stabile (ad esempio `10.0.0.0`) non è la versione patch del pacchetto: distinguere sempre assembly, product version e origine del file. I moduli possono distribuire loader/DLL con versioni diverse da quella del modulo.

Fonti primarie consultate il 27 ottobre 2026:

- [.NET 10: SDK 10.0.403 e runtime 10.0.14](https://dotnet.microsoft.com/en-us/download/dotnet/10.0), [aggiornamento di sicurezza OOB](https://devblogs.microsoft.com/dotnet/dotnet-october-2026-oob-security-updates/).
- [PowerShell 7.6.9](https://github.com/PowerShell/PowerShell/releases/tag/v7.6.9), [manifest NuGet esatto e dipendenze minime](https://api.nuget.org/v3-flatcontainer/microsoft.powershell.sdk/7.6.9/microsoft.powershell.sdk.nuspec).
- [EXO 3.10.1](https://www.powershellgallery.com/packages/ExchangeOnlineManagement/3.10.1), [requisiti EXO/PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps), [Graph Authentication 2.41.2](https://www.powershellgallery.com/packages/Microsoft.Graph.Authentication/2.41.2).
- [xUnit 4: cambiamenti e runner](https://xunit.net/releases/v3/4.0.0), [correzioni 4.0.2](https://xunit.net/releases/v3/4.0.2), [StaFact 4.0.24](https://www.nuget.org/packages/Xunit.StaFact/4.0.24), [Test SDK 18.11.0](https://www.nuget.org/packages/Microsoft.NET.Test.Sdk/18.11.0), [coverlet 10.1.0](https://www.nuget.org/packages/coverlet.collector/10.1.0).
- [Gitleaks 8.30.1](https://github.com/gitleaks/gitleaks/releases/tag/v8.30.1), [checksum ufficiali](https://github.com/gitleaks/gitleaks/releases/download/v8.30.1/gitleaks_8.30.1_checksums.txt), [mitigazione script injection GitHub](https://docs.github.com/en/actions/reference/security/secure-use#good-practices-for-mitigating-script-injection-attacks).

Gli aggiornamenti riducono lo scostamento dalla manutenzione corrente. L'audit NuGet era già privo di advisory noti prima del lotto: non si attribuisce a ogni aggiornamento la correzione di una vulnerabilità dimostrata.

## Verifiche ed evidenze

Baseline precedente: 887/887 test. Finale: **919/919 test Release, zero fallimenti o ignorati**, 32 casi aggiunti; build Release locked senza errori; audit NuGet su quattro progetti e Gitleaks 8.30.1 passati; 33 script senza errori parser/analyzer e vincoli architetturali rispettati. Risultati e comandi integrali nel dossier SecondBrain `history/2026-10-27-application-audit/`, con copie dei report e manifest SHA-256. Il report macchina `verification.json` distingue controlli statici, test e limiti.

Comandi principali, dalla radice:

```powershell
dotnet test --nologo -c Release --no-restore --artifacts-path artifacts/build
pwsh -NoProfile -File scripts/build.ps1 -Configuration Release -RuntimeIdentifier win-x64 -Clean:$false -LockedMode
pwsh -NoProfile -File build/assert-no-vulnerable-packages.ps1
pwsh -NoProfile -File build/run-secret-scan.ps1
pwsh -NoProfile -File build/assert-architecture-constraints.ps1
```

Test mirati includono certificati unico/assente/ambiguo/scaduto/futuro/senza chiave, subject con apice, rinnovo ed esplicito thumbprint; binding senza parametri inesistenti; timeout/cancellazione in lock/write/response/importazione/HTTP/processi; saturazione della capacità. CMS/XML usano un certificato effimero in memoria e verificano anche la modifica del contenuto firmato.

La discovery confronta 882 nomi mostrati prima e 914 dopo: nessuna voce precedente persa. Cinque casi vengono espansi soltanto durante l'esecuzione, perciò il numero mostrato da `--list-tests` differisce di cinque dal totale eseguito in entrambe le versioni. Coverage: 140 test passati con lo stesso filtro CI, 24,15% complessivo sui pacchetti Worker/Contracts (3832/15867 righe), sopra la soglia invariata del 20%, senza migrazione a MTP. Non è la copertura dell'intera applicazione.

Carichi sintetici dell'ultima suite: mapping di 10.000 record circa 24 ms, raccolta da 10.000 a 5.000 righe circa 5 ms, 1.000 log con riepilogo circa 211 ms, XLSX da 10.000 righe circa 282 ms con formule trattate come testo. Tempi osservati conservati negli output TRX, senza soglie temporali fragili nei test. Non dimostrano latenza di rete o reattività di griglie WPF popolate.

PSScriptAnalyzer: 147 warning analizzati per regola, senza soppressioni aggiunte: 100 `WriteHost` (output operativo), 22 nomi plurali, 8 switch con default esplicito, 7 `ShouldProcess` (incluse factory e helper di harness), 5 verbi, 4 parametri segnalati inutilizzati (inclusi riferimenti in scriptblock), 1 password stringa della firma. I warning stilistici non diventano automaticamente task; l'ultimo contribuisce al residuo confermato seguente.

## Residuo confermato e proposte separate

**P2, REV-20261027-02 — firma PFX:** `build/sign-artifacts.ps1` passa la password a SignTool con `/p`. Il valore entra negli argomenti del processo nativo e può essere letto da soggetti locali con accesso ai metadati di quel processo: [CWE-214](https://cwe.mitre.org/data/definitions/214.html). È confermato il percorso nel codice, non una compromissione né una fuga osservata di credenziali reali.

Proposta fuori dal collaudo attuale: adottare una firma senza password negli argomenti, con backend di firma o certificato su store isolato; verificare chiavi/ACL, firma, timestamp e cleanup con credenziali di test. Non sostituire silenziosamente il backend né importare certificati nello store personale. Task bloccato sulla scelta/ambiente di firma; nessuna chiusura basata sul solo warning.

**P2, REV-20261027-03 — dipendenza inclusa in EXO:** il pacchetto ufficiale stabile EXO 3.10.1 include `netCore/System.Security.Cryptography.Pkcs.dll` prodotto 10.0.0 e lo dichiara nel proprio `.deps.json`. Hash del file installato confrontato con il pacchetto ufficiale; versione compresa nei range di [GHSA-846f-f764-76mx](https://github.com/dotnet/runtime/security/advisories/GHSA-846f-f764-76mx) e [GHSA-2fmv-rv2c-59x7](https://github.com/dotnet/runtime/security/advisories/GHSA-2fmv-rv2c-59x7). Il feed NuGet controllato per sei famiglie note fra le DLL inventariate conferma questi due match; non è una scansione esaustiva di tutti i binari dei moduli.

La presenza del componente affetto è confermata, il suo utilizzo vulnerabile dall'applicazione non è dimostrato: il probe dell'host carica Pkcs 10.0.14. Serve qualificare risoluzione e contesti di caricamento nei percorsi EXO di autenticazione, oppure adottare un aggiornamento stabile supportato del fornitore. Non alterare DLL nel modulo personale, non adottare preview e non dedurre dall'audit NuGet della soluzione che il pacchetto EXO sia bonificato. Il residuo resta separato e bloccato sulla qualificazione/soluzione supportata del modulo esterno.

Aggiornamenti non applicati perché richiedono qualificazione aggiuntiva:

- Actions: riferimenti correnti v4; candidate ufficiali [checkout 7.0.1](https://github.com/actions/checkout/releases/tag/v7.0.1), [setup-dotnet 6.0.0](https://github.com/actions/setup-dotnet/releases/tag/v6.0.0), [upload-artifact 7.0.2](https://github.com/actions/upload-artifact/releases/tag/v7.0.2). Qualificare runtime del runner, cache, upload e release su CI senza pubblicazione prima di cambiare i pin. Nessuna CI remota avviata.
- Inno Setup: CI ancora 6.7.1; [6.7.3 e ramo 7 ufficiali](https://jrsoftware.org/isdl.php). Prima della modifica qualificare compilazione, riproducibilità e install/uninstall su VM isolata. Nessuna installazione di strumenti o prova installer sulla postazione.
- .NET 11 e PowerShell 7.7 prerelease esclusi; nessun motivo di compatibilità emerso per cambiare le major runtime stabili.

UX-08 e CLEAN-01 restano aperti con i vincoli esistenti. Scratch storico NuGet e artefatti preesistenti preservati; i soli temporanei posseduti da questo audit sono rimossi dopo aver salvato le evidenze.

Non verificati: autenticazione e operazioni tenant reali, comportamento completo di Exchange/Graph/Purview, accessibilità/DPI/high contrast/Narrator, prestazioni UI reali, installer/signing con credenziali reali, CI remota, compatibilità delle major Actions/Inno e audit esaustivo di tutte le DLL incluse nei moduli. La patch del runtime della macchina non aggiorna eventuali installer self-contained già prodotti: serve ricostruzione e qualificazione separata degli artefatti distribuiti.
