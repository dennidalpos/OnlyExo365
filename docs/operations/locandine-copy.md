# Materiale di Comunicazione e Copy per Locandine — OnlyExo365

Revisione: 1.0 — 2026-10-22  
Scopo: Raccolta di testi pronti all'uso per locandine, roll-up fieristici, flyer promozionali e schede informative per utenti finali, amministratori di sistema e responsabili IT.

---

## Indice Formati

1. [Format A — Roll-up Verticale per Eventi e Fiere (85x200 cm)](#format-a--roll-up-verticale-85x200-cm)
2. [Format B — Locandina A3/A4 per Bacheca IT e Uffici](#format-b--locandina-a3a4-per-bacheca-it)
3. [Format C — Flyer Pieghevole / Scheda Prodotto Fronte-Retro](#format-c--flyer-pieghevole-fronte-retro)
4. [Format D — Banner Grafici e Social Card (16:9 / 1:1)](#format-d--banner-grafici-e-social-card)
5. [Guida alle Immagini Collegate](#guida-alle-immagini-collegate)

---

## Format A — Roll-up Verticale (85x200 cm)

> **Obiettivo**: Catturare l'attenzione a 3-5 metri di distanza, comunicando in 5 secondi il valore dell'applicazione: Exchange Online sotto controllo, senza complessità da riga di comando.

### Header (In alto)
- **Logo**: OnlyExo365 (`src/OnlyExo365.Shell/Assets/Generated/AppIcon.png`)
- **Headline**: **Exchange Online, finalmente semplice.**
- **Sottotitolo**: Tutta la potenza di Microsoft 365 sul tuo desktop Windows, senza script complicati.

### Body (Al centro — 4 Punti di Forza con Icone)
* **⚡ Cruscotto Istantaneo**  
  Stato del tenant, quote caselle e licenze a colpo d'occhio. Zero attese.
* **🛡️ Sicurezza Senza Errori**  
  Anti-spam, anti-phishing, DKIM e regole di posta con salvaguardia delle bozze: vedi prima di applicare.
* **🔍 Tracciamento Email in 3 Clic**  
  Verifica se un messaggio è arrivato, bloccato o inoltrato senza impazzire nei log di sistema.
* **💼 Conformità e Audit Immediati**  
  Ricerche Purview integrate ed esportazione report con un solo click.

### Visual Centrale
- Riquadro con mockup o screenshot della schermata principale: `docs/assets/screenshots/01-dashboard.png`.

### Footer (In basso)
- **Call to Action**: **Scarica la versione desktop x64 per Windows.**
- **Requisiti**: Windows 10/11 x64 • .NET 10 • Modulo Exchange Online
- **Badge**: 100% Sicuro • Nessuna credenziale in chiaro • Esecuzione isolata

---

## Format B — Locandina A3/A4 per Bacheca IT

> **Obiettivo**: Esposizione interna negli uffici helpdesk, dipartimenti IT e sale server per guidare i tecnici di supporto nella gestione quotidiana del tenant.

```markdown
┌────────────────────────────────────────────────────────────────────────┐
│                              OnlyExo365                                │
│          La console desktop che velocizza il tuo Helpdesk M365         │
└────────────────────────────────────────────────────────────────────────┘

Basta comandi PowerShell complessi per le operazioni di tutti i giorni.
OnlyExo365 trasforma le attività su Exchange Online in pochi click guidati e sicuri.

 COSA PUOI FARE SUBITO:

 1. GESTIRE CASSETTE POSTALI E PERMESSI
    Crea caselle, imposta risposte automatiche fuori sede e gestisci deleghe
    (Invia come, Accesso completo) con anteprima delle modifiche.
    Schermata: docs/assets/screenshots/02-cassette-postali.png

 2. RISOLVERE PROBLEMI DI RECAPITO EMAIL
    Un utente lamenta una mail non ricevuta? Inserisci mittente o destinatario
    e scopri subito stato, data e motivazione del blocco.
    Schermata: docs/assets/screenshots/03-tracciamento-messaggi.png

 3. PROTEGGERE IL DOMINIO E LA POSTA
    Verifica lo stato di DKIM, aggiorna le liste di quarantena e regola
    i filtri anti-spam con regole chiare e trasparenti.
    Schermata: docs/assets/screenshots/04-sicurezza-posta.png

 4. VERIFICA E AUDIT DI CONFORMITÀ
    Monitora le attività sul tenant ed esporta i log di controllo per gli audit.
    Schermata: docs/assets/screenshots/05-conformita-purview.png

 ────────────────────────────────────────────────────────────────────────
 SICURO PER DESIGN:
 • Nessuna password salvata: autenticazione diretta Microsoft 365.
 • Motore isolato: l'interfaccia resta sempre fluida e reattiva.
 • Bozza sicura: controlli prima, applichi solo quando sei sicuro.
 ────────────────────────────────────────────────────────────────────────
 Disponibile per postazioni Windows 64-bit aziendali.
```

---

## Format C — Flyer Pieghevole Fronte-Retro

### Lato Fronte (Copertina)
- **Titolo ad impatto**: **Dì addio allo stress dell'amministrazione Exchange.**
- **Visual**: Schermata completa della Dashboard (`docs/assets/screenshots/01-dashboard.png`) inserita in una cornice monitor moderna.
- **Payoff**: "La console nativa Windows che rende Microsoft Exchange Online intuitivo, veloce e a prova di errore."
- **Tagline**: Per amministratori IT, tecnici di supporto e responsabili dei servizi cloud.

### Lato Retro (Dettaglio Operativo & Vantaggi)
- **Le 5 funzioni che userai ogni giorno**:
  1. **Dashboard di controllo**: Panoramica della salute del tenant e stato dei servizi in tempo reale.
  2. **Inventario Utenti e Caselle**: Ricerca istantanea, quote disco e modifica impostazioni senza console web lente.
  3. **Traccia Messaggi Express**: Diagnostica dei flussi email in pochi secondi, con esportazione diretta in Excel.
  4. **Regole e Connettori**: Gestione visuale delle regole di trasporto (`docs/assets/screenshots/06-flusso-posta.png`) per indirizzare correttamente la posta aziendale.
  5. **Diagnostica e Strumenti**: Verifica automatica della compatibilità dei moduli PowerShell e console trasparente per audit (`docs/assets/screenshots/07-strumenti-diagnostica.png`).
- **I 3 pilastri tecnici**:
  - **Zero rischi**: le modifiche passano da una fase di bozza e convalida prima dell'invio.
  - **Velocità nativa**: applicazione compilata in C# / .NET x64 per massime prestazioni su Windows.
  - **Privacy totale**: niente cloud di terze parti, il client comunica direttamente con il tuo tenant Microsoft.

---

## Format D — Banner Grafici e Social Card

### Variante 1: Orientamento Orizzontale (1200x627 px — LinkedIn / Twitter Card)
- **Testo principale**: "Gestire Exchange Online non richiede più ore di script."
- **Testo secondario**: "OnlyExo365: la console desktop intuitiva per l'amministrazione Microsoft 365."
- **Pulsante simulato (CTA)**: "Scopri le funzionalità desktop"
- **Immagine di sfondo**: `docs/assets/screenshots/01-dashboard.png` con overlay gradiente scuro.

### Variante 2: Orientamento Quadrato (1080x1080 px — Instagram / Feed Post)
- **Titolo**: **OnlyExo365**
- **Sottotitolo**: **Amministrazione Exchange Online a portata di click.**
- **Icone/Punti chiave**:
  - ✔️ Ricerca e gestione caselle
  - ✔️ Tracciamento email in tempo reale
  - ✔️ Controllo quote e licenze
  - ✔️ Sicurezza DKIM & anti-spam
- **Immagine abbinata**: Carosello con le schermate `02-cassette-postali.png`, `03-tracciamento-messaggi.png` e `04-sicurezza-posta.png`.

---

## Guida alle Immagini Collegate

Tutti gli asset grafici sono generati in alta risoluzione (1280x880 pixel) in formato uniforme PNG:

| Schermata | File Immagine | Sezione Descritta |
|---|---|---|
| Prima Pagina (Dashboard) | [`01-dashboard.png`](file:///d:/GITHUB/OnlyExo365/docs/assets/screenshots/01-dashboard.png) | Panoramica stato tenant e indicatori di connessione |
| Cassette Postali | [`02-cassette-postali.png`](file:///d:/GITHUB/OnlyExo365/docs/assets/screenshots/02-cassette-postali.png) | Inventario, quote, creazione caselle e permessi |
| Traccia Messaggi | [`03-tracciamento-messaggi.png`](file:///d:/GITHUB/OnlyExo365/docs/assets/screenshots/03-tracciamento-messaggi.png) | Ricerca recapito email, filtri data e dettagli consegna |
| Sicurezza Posta | [`04-sicurezza-posta.png`](file:///d:/GITHUB/OnlyExo365/docs/assets/screenshots/04-sicurezza-posta.png) | Baseline DKIM, anti-spam, anti-phishing e quarantena |
| Conformità Purview | [`05-conformita-purview.png`](file:///d:/GITHUB/OnlyExo365/docs/assets/screenshots/05-conformita-purview.png) | Log di controllo unificato e audit ricerche |
| Flusso di Posta | [`06-flusso-posta.png`](file:///d:/GITHUB/OnlyExo365/docs/assets/screenshots/06-flusso-posta.png) | Regole di trasporto, connettori e domini accettati |
| Strumenti & Diagnostica | [`07-strumenti-diagnostica.png`](file:///d:/GITHUB/OnlyExo365/docs/assets/screenshots/07-strumenti-diagnostica.png) | Console worker trasparente, catalogo licenze e requisiti |
