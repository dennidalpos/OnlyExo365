# Guida visiva ai percorsi utente

Questa guida illustra i flussi operativi dell'applicazione attraverso diagrammi visivi dedicati per ciascuna funzionalità principale. Ogni percorso evidenzia il punto di ingresso dell'utente, le azioni disponibili, le scelte decisionali, i risultati previsti e le procedure di recupero in caso di errore.

---

## Indice dei percorsi

1. [Connessione e Autenticazione al Tenant](#1-connessione-e-autenticazione-al-tenant)
2. [Dashboard e Panoramica Generale](#2-dashboard-e-panoramica-generale)
3. [Gestione Cassette Postali e Dettagli Utente](#3-gestione-cassette-postali-e-dettagli-utente)
4. [Tracciamento dei Messaggi (Message Trace)](#4-tracciamento-dei-messaggi-message-trace)
5. [Politiche di Sicurezza della Posta (Mail Security)](#5-politiche-di-sicurezza-della-posta-mail-security)
6. [Conformità e Ricerca Purview (Compliance & eDiscovery)](#6-conformit%C3%A0-e-ricerca-purview-compliance--ediscovery)
7. [Regole e Flusso di Posta (Mail Flow)](#7-regole-e-flusso-di-posta-mail-flow)
8. [Strumenti, Prerequisiti e Diagnostica (Tools)](#8-strumenti-prerequisiti-e-diagnostica-tools)

---

## 1. Connessione e Autenticazione al Tenant

Il percorso iniziale permette all'operatore di autenticarsi verso il tenant Microsoft 365, inizializzare il motore di gestione e sbloccare le funzionalità dell'applicazione.

```mermaid
flowchart TD
    Start(["Avvio applicazione"]) --> CheckPolicy{"Connessioni abilitate da policy?"}
    
    CheckPolicy -->|"No (Disabilitate da configurazione)"| BlockedBanner["Visualizzazione avviso di blocco in evidenza<br/>Pulsante Connetti disattivato"]
    BlockedBanner --> FixPolicy["Operatore verifica criteri locali o variabili d'ambiente"]
    
    CheckPolicy -->|"Sì (Abilitate)"| ClickConnect["Pulsante 'Connetti' nella barra superiore o Strumenti"]
    ClickConnect --> ChooseAuth{"Modalità di autenticazione configurata"}
    
    ChooseAuth -->|"Accesso interattivo"| BrowserLogin["Apertura browser / finestra di login Microsoft 365<br/>Immissione credenziali e approvazione MFA"]
    ChooseAuth -->|"Codice dispositivo"| DeviceCodeLogin["Visualizzazione codice e indirizzo web<br/>Autenticazione da dispositivo secondario"]
    ChooseAuth -->|"Certificato applicativo"| CertLogin["Verifica certificato locale installato nel sistema<br/>Convalida impronta digitale e identificativo tenant"]
    ChooseAuth -->|"Identità gestita"| ManagedIdLogin["Rilevamento identità assegnata all'ambiente host"]
    
    BrowserLogin --> ConnectOutcome{"Esito autenticazione"}
    DeviceCodeLogin --> ConnectOutcome
    CertLogin --> ConnectOutcome
    ManagedIdLogin --> ConnectOutcome
    
    ConnectOutcome -->|"Successo"| CheckModules{"Moduli approvati disponibili?"}
    ConnectOutcome -->|"Credenziali errate o MFA annullata"| LoginError["Notifica di mancata autenticazione<br/>Dettagli di supporto visualizzati"]
    LoginError --> ClickConnect
    
    CheckModules -->|"Sì"| SessionActive["Stato: Connesso<br/>Visualizzazione utente e tenant nella barra in alto<br/>Moduli operativi sbloccati nella navigazione"]
    CheckModules -->|"No (Modulo mancante o versione errata)"| ModuleError["Notifica: modulo Exchange Online non pronto<br/>Rimando alla scheda Strumenti"]
    ModuleError --> OpenTools["Accesso a Strumenti per installare i componenti approvati"]
```

---

## 2. Dashboard e Panoramica Generale

Pannello di controllo che sintetizza le informazioni chiave del tenant e offre accessi rapidi alle sezioni operative.

```mermaid
flowchart TD
    EnterDashboard["Accesso alla scheda 'Dashboard'"] --> CheckConn{"Connessione attiva?"}
    
    CheckConn -->|"No"| PromptConnect["Invito a connettersi al tenant<br/>Metriche e statistiche non disponibili"]
    PromptConnect --> GoConnect["Click su 'Connetti' per iniziare la sessione"]
    
    CheckConn -->|"Sì"| LoadData["Caricamento automatico metriche e quote"]
    LoadData --> DataOutcome{"Esito consultazione tenant"}
    
    DataOutcome -->|"Successo"| ShowStats["Visualizzazione riepilogo cassette postali, licenze e spazio<br/>Visualizzazione avvisi o scadenze rilevate"]
    ShowStats --> UserAction{"Azione operatore"}
    
    UserAction -->|"Click su collegamento rapido"| NavigateTarget["Apertura diretta del modulo corrispondente<br/>(es. Cassette postali, Sicurezza, Flusso)"]
    UserAction -->|"Click 'Aggiorna'"| LoadData
    
    DataOutcome -->|"Timeout o indisponibilità servizio"| ShowWarning["Avviso di caricamento parziale o fallito<br/>Pulsante 'Riprova' disponibile"]
    ShowWarning --> LoadData
```

---

## 3. Gestione Cassette Postali e Dettagli Utente

Ricerca, ispezione delle impostazioni e modifica protetta delle proprietà delle cassette postali con salvaguardia delle bozze non salvate.

```mermaid
flowchart TD
    EnterMailboxes["Accesso alla scheda 'Cassette postali'"] --> SearchBox["Filtro elenco con ricerca testo<br/>(Nome, cognome, indirizzo email o UPN)"]
    SearchBox --> SelectMailbox["Selezione di una cassetta postale dalla griglia"]
    SelectMailbox --> ViewDetails["Apertura pannello dettagli utente:<br/>Informazioni account, licenze, quote e risposte automatiche"]
    
    ViewDetails --> InspectTabs{"Scelta ambito da gestire"}
    
    InspectTabs -->|"Impostazioni generali & Quote"| EditGeneral["Modifica nome visualizzato, limiti di spazio o inoltro"]
    InspectTabs -->|"Risposte automatiche"| EditAutoReply["Attivazione/disattivazione e modifica messaggi di assenza"]
    InspectTabs -->|"Criteri di conservazione"| EditRetention["Scelta politica di conservazione (Retention Policy)"]
    InspectTabs -->|"Permessi di accesso"| EditPermissions["Aggiunta o rimozione deleghe (Full Access, Send As, Send on Behalf)"]
    
    EditGeneral --> ChangeDetected["Rilevamento modifiche in sospeso (Bozza attiva)"]
    EditAutoReply --> ChangeDetected
    EditRetention --> ChangeDetected
    EditPermissions --> ChangeDetected
    
    ChangeDetected --> UserChoice{"Decisione dell'operatore"}
    
    UserChoice -->|"Salva modifiche"| ConfirmDialog["Finestra modale di conferma con riepilogo impatto"]
    ConfirmDialog -->|"Conferma"| ApplyToCloud["Invio richiesta di aggiornamento a Microsoft 365"]
    ConfirmDialog -->|"Annulla"| EditGeneral
    
    ApplyToCloud --> CloudResult{"Esito aggiornamento"}
    CloudResult -->|"Completato con successo"| SuccessUpdate["Stato salvato<br/>Bozza azzerata e dati sincronizzati"]
    CloudResult -->|"Errore o conflitto"| ShowError["Notifica di errore descrittiva<br/>Bozza preservata per permettere correzioni"]
    ShowError --> EditGeneral
    
    UserChoice -->|"Scarta modifiche"| DiscardDraft["Ripristino dei valori originali della cassetta postale"]
    
    UserChoice -->|"Cambio scheda o selezione utente"| DirtyPrompt{"Avviso bozza non salvata"}
    DirtyPrompt -->|"Salva ed esci"| ConfirmDialog
    DirtyPrompt -->|"Scarta ed esci"| DiscardDraft
    DirtyPrompt -->|"Rimani qui"| EditGeneral
```

---

## 4. Tracciamento dei Messaggi (Message Trace)

Strumento per localizzare messaggi di posta, verificare lo stato di consegna e diagnosticare problemi di recapito.

```mermaid
flowchart TD
    EnterTrace["Accesso alla scheda 'Tracciamento messaggi'"] --> SetFilters["Compilazione criteri di ricerca:<br/>- Indirizzo mittente (opzionale)<br/>- Indirizzo destinatario (opzionale)<br/>- Data inizio e data fine<br/>- Dimensione pagina"]
    
    SetFilters --> ClickSearch["Click su 'Cerca messaggi'"]
    ClickSearch --> ValidateDates{"Verifica intervallo temporale"}
    
    ValidateDates -->|"Intervallo > 10 giorni o date future o > 90 giorni fa"| DateError["Avviso a video: intervallo non valido<br/>(Massimo 10 giorni consecutivi negli ultimi 90 giorni)"]
    DateError --> SetFilters
    
    ValidateDates -->|"Intervallo valido"| RunQuery["Esecuzione ricerca nel registro messaggi di Exchange"]
    RunQuery --> QueryResult{"Risultato ricerca"}
    
    QueryResult -->|"Messaggi trovati"| ShowGrid["Visualizzazione elenco messaggi:<br/>Mittente, destinatario, oggetto, data/ora, stato consegna"]
    QueryResult -->|"Nessun messaggio trovato"| NoResults["Notifica: nessun messaggio corrispondente ai criteri"]
    NoResults --> SetFilters
    QueryResult -->|"Timeout o servizio occupato"| QueryError["Avviso di interruzione o ritardo<br/>Suggerimento di restringere la finestra temporale"]
    QueryError --> SetFilters
    
    ShowGrid --> PageActions{"Azioni sui risultati"}
    
    PageActions -->|"Filtro rapido stato"| FilterStatus["Selezione filtro: Tutti, Consegnati, Bloccati, In attesa"]
    FilterStatus --> ShowGrid
    
    PageActions -->|"Navigazione pagine"| ChangePage["Click su 'Pagina successiva' o 'Pagina precedente'"]
    ChangePage --> RunQuery
    
    PageActions -->|"Dettaglio messaggio"| SelectMessage["Selezione messaggio e click 'Dettagli traccia'"]
    SelectMessage --> ViewEvents["Visualizzazione eventi completi di routing ed eventuali motivi di blocco"]
    
    PageActions -->|"Esporta"| ExportExcel["Click su 'Esporta in Excel'<br/>Salvataggio file .xlsx sul disco locale"]
```

---

## 5. Politiche di Sicurezza della Posta (Mail Security)

Gestione delle difese perimetrali di Exchange Online: criteri anti-spam, anti-phishing, anti-malware, quarantena e chiavi DKIM.

```mermaid
flowchart TD
    EnterSecurity["Accesso alla scheda 'Sicurezza posta'"] --> SelectPolicyCategory{"Selezione categoria di protezione"}
    
    SelectPolicyCategory -->|"Anti-Spam"| LoadSpam["Consultazione filtri spam in entrata/uscita"]
    SelectPolicyCategory -->|"Anti-Phishing"| LoadPhish["Consultazione protezione spoofing e impersonificazione"]
    SelectPolicyCategory -->|"Anti-Malware"| LoadMalware["Consultazione filtri malware e tipi di allegati bloccati"]
    SelectPolicyCategory -->|"Quarantena"| LoadQuarantine["Consultazione criteri di notifica e rilascio quarantena"]
    SelectPolicyCategory -->|"DKIM"| LoadDkim["Consultazione chiavi e firme per i domini aziendali"]
    
    LoadSpam --> EditPolicy["Modifica impostazioni della policy selezionata"]
    LoadPhish --> EditPolicy
    LoadMalware --> EditPolicy
    LoadQuarantine --> EditPolicy
    LoadDkim --> EditPolicy
    
    EditPolicy --> CheckDirty["Rilevamento modifiche pendenti"]
    CheckDirty --> TrySave["Click su 'Salva' della policy corrispondente"]
    
    TrySave --> CheckOtherDrafts{"Ci sono modifiche non salvate in altre categorie?"}
    
    CheckOtherDrafts -->|"Sì (Bozze multiple aperte)"| MultiDraftDialog{"Dialogo di conferma:<br/>Il salvataggio confermerà questa policy e scarterà le altre bozze non salvate"}
    MultiDraftDialog -->|"Salva e scarta le altre"| SubmitCloud["Invio aggiornamento a Microsoft 365"]
    MultiDraftDialog -->|"Rimani qui per completare"| EditPolicy
    
    CheckOtherDrafts -->|"No (Solo la policy corrente)"| NormalConfirm["Conferma salvataggio criteri di sicurezza"]
    NormalConfirm --> SubmitCloud
    
    SubmitCloud --> SaveResult{"Esito applicazione"}
    SaveResult -->|"Completato"| PolicyUpdated["Policy aggiornata e attiva nel tenant<br/>Contrassegni di modifica azzerati"]
    SaveResult -->|"Errore permessi o formato"| PolicyError["Notifica di errore con motivazione fornita dal servizio<br/>Bozza mantenuta per correzioni"]
    PolicyError --> EditPolicy
```

---

## 6. Conformità e Ricerca Purview (Compliance & eDiscovery)

Gestione delle ricerche di conformità legale, analisi dei log di audit unificati e applicazione di blocchi protettivi (Hold).

```mermaid
flowchart TD
    EnterCompliance["Accesso alla scheda 'Conformità'"] --> CheckComplianceSession{"Sessione di conformità disponibile?"}
    
    CheckComplianceSession -->|"No (Autenticazione non delegata o sessione assente)"| SessionWarning["Avviso: Purview richiede autenticazione interattiva delegata<br/>Modalità automatica o certificato non supportate per eDiscovery"]
    SessionWarning --> ReconnectInteractively["Riconnessione tramite modalità Interattiva con account abilitato"]
    
    CheckComplianceSession -->|"Sì"| ComplianceWorkspace["Spazio di lavoro Conformità attivo:<br/>Ricerche esistenti, casi di conformità e stato azioni"]
    
    ComplianceWorkspace --> UserActivity{"Scelta attività"}
    
    UserActivity -->|"Ricerca Log di Audit"| AuditSearchForm["Impostazione intervallo date, attività utente o tipo evento"]
    AuditSearchForm --> RunAudit["Avvio interrogazione registro di audit unificato"]
    RunAudit --> ShowAuditResults["Visualizzazione risultati con data, utente, IP e operazione eseguita"]
    
    UserActivity -->|"Nuova Ricerca eDiscovery"| CreateSearch["Compilazione nome ricerca, posizioni Exchange da includere e query KQL"]
    CreateSearch --> SaveSearch["Click su 'Crea ricerca'"]
    SaveSearch --> StartSearchChoice{"Avviare subito la ricerca?"}
    StartSearchChoice -->|"Sì"| ExecuteSearch["Esecuzione ricerca nei dati del tenant"]
    StartSearchChoice -->|"No"| SearchSaved["Ricerca memorizzata per esecuzione futura"]
    
    UserActivity -->|"Gestione ricerca esistente"| SelectExistingSearch["Selezione ricerca dalla lista"]
    SelectExistingSearch --> SearchActions{"Azione sulla ricerca"}
    
    SearchActions -->|"Avvia / Riavvia"| ExecuteSearch
    ExecuteSearch --> SearchCompleted["Risultati disponibili: volume stimato e messaggi individuati"]
    
    SearchActions -->|"Crea blocco (Hold)"| HoldPrompt["Inserimento nome blocco e conferma associazione al caso"]
    HoldPrompt --> ConfirmHold["Conferma applicazione blocco legale di conservazione"]
    ConfirmHold --> HoldActive["Blocco attivo: i contenuti corrispondenti non possono essere eliminati"]
    
    SearchActions -->|"Eliminazione mirata (Purge)"| PurgePrompt["Finestra modale di sicurezza: confermare eliminazione definitiva dei messaggi individuati"]
    PurgePrompt -->|"Conferma con riepilogo"| ExecutePurge["Esecuzione rimozione protetta dei messaggi da parte del tenant"]
    PurgePrompt -->|"Annulla"| SelectExistingSearch
```

---

## 7. Regole e Flusso di Posta (Mail Flow)

Configurazione dell'infrastruttura di routing messaggi: regole di trasporto globali, connettori aziendali, domini accettati e relazioni organizzative.

```mermaid
flowchart TD
    EnterMailFlow["Accesso alla scheda 'Flusso di posta'"] --> SelectSection{"Selezione ambito di instradamento"}
    
    SelectSection -->|"Regole di trasporto"| ManageRules["Consultazione regole di trasporto globali del tenant"]
    SelectSection -->|"Connettori"| ManageConnectors["Consultazione connettori in entrata e uscita"]
    SelectSection -->|"Domini accettati"| ManageDomains["Consultazione domini autorizzati (Authoritative o Internal Relay)"]
    SelectSection -->|"Relazioni organizzative"| ManageOrgs["Consultazione federazioni e condivisione calendario con tenant partner"]
    SelectSection -->|"Elenchi indirizzi"| ManageAddressLists["Consultazione rubriche e policy elenchi aziendali"]
    
    ManageRules --> RuleAction{"Operazione su regola"}
    
    RuleAction -->|"Nuova regola"| DefineNewRule["Definizione nome, priorità, condizioni, azioni ed eccezioni"]
    RuleAction -->|"Modifica esistente"| EditRule["Aggiornamento parametri della regola selezionata"]
    RuleAction -->|"Attiva / Disattiva"| ToggleRule["Commutazione stato abilitata/disabilitata"]
    RuleAction -->|"Elimina"| DeleteRulePrompt["Conferma eliminazione regola"]
    
    DefineNewRule --> ValidateRule{"Verifica validità parametri"}
    EditRule --> ValidateRule
    ToggleRule --> ValidateRule
    DeleteRulePrompt --> ValidateRule
    
    ValidateRule -->|"Campi obbligatori completi"| SubmitRuleChanges["Invio configurazione a Exchange Online"]
    ValidateRule -->|"Parametri incompleti o contrastanti"| RuleInputError["Evidenza campi errati nel modulo"]
    RuleInputError --> DefineNewRule
    
    SubmitRuleChanges --> RuleOutcome{"Esito aggiornamento"}
    RuleOutcome -->|"Successo"| RulesRefreshed["Regola applicata e lista riordinata per priorità"]
    RuleOutcome -->|"Errore configurazione remota"| RuleServerError["Segnalazione errore con dettaglio fornito dal server"]
    RuleServerError --> EditRule
```

---

## 8. Strumenti, Prerequisiti e Diagnostica (Tools)

Pannello tecnico per la verifica dell'ambiente di esecuzione, l'installazione dei moduli ufficiali necessari, il monitoraggio diagnostico e l'aggiornamento del catalogo licenze.

```mermaid
flowchart TD
    EnterTools["Accesso alla scheda 'Strumenti'"] --> AutoCheck["Esecuzione automatica verifica prerequisiti"]
    
    AutoCheck --> DisplayPrereqs["Riepilogo stato componenti:<br/>- Presenza ed edizione di PowerShell (consigliata 7.6+)<br/>- Modulo ExchangeOnlineManagement approvato<br/>- Moduli Microsoft.Graph approvati<br/>- Stato aggiornamento catalogo licenze M365"]
    
    DisplayPrereqs --> PrereqStatus{"Tutti i componenti sono pronti?"}
    
    PrereqStatus -->|"Sì (Stato ottimale)"| ReadyState["Tutte le funzioni del tenant sono operative"]
    PrereqStatus -->|"No (Moduli mancanti o non approvati)"| ShowInstallActions["Pulsanti 'Installa' abilitati accanto ai moduli mancanti"]
    
    ShowInstallActions --> ClickInstall["Click su 'Installa modulo'"]
    ClickInstall --> RunModuleSetup["Download e installazione della versione approvata da PSGallery"]
    
    RunModuleSetup --> InstallOutcome{"Esito installazione"}
    InstallOutcome -->|"Successo"| RefreshStatus["Riconvalida automatica prerequisiti: modulo pronto"]
    InstallOutcome -->|"Errore connessione o permessi"| ManualGuidance["Visualizzazione comando alternativo da eseguire manualmente nel terminale"]
    ManualGuidance --> DisplayPrereqs
    
    ReadyState --> AdvancedTools{"Strumenti avanzati operatore"}
    DisplayPrereqs --> AdvancedTools
    
    AdvancedTools -->|"Controllo console diagnostica"| ToggleConsole["Spunta 'Mostra console diagnostica in tempo reale'"]
    ToggleConsole --> ConsoleVisibility{"Azione selezionata"}
    ConsoleVisibility -->|"Mostra"| WindowShown["Apertura console Worker con replay storico eventi"]
    ConsoleVisibility -->|"Nascondi"| WindowHidden["Console nascosta in background senza arrestare il motore"]
    
    AdvancedTools -->|"Catalogo licenze M365"| RefreshCatalog["Click su 'Aggiorna catalogo licenze ora'"]
    RefreshCatalog --> DownloadCatalog["Download e validazione elenco ufficiale piani Microsoft"]
    DownloadCatalog --> CatalogUpdated["Nomi commerciali delle licenze sincronizzati"]
    
    AdvancedTools -->|"Esportazione log"| ExportLogs["Click su 'Esporta archivio log diagnostici'<br/>Salvataggio file .zip contenente log di interfaccia e motore"]
```
