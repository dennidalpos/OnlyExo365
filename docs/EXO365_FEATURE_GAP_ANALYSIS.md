# Analisi gap funzionali OnlyExo365 vs gestione Exchange Online (baseline Microsoft)

## 1) Stato attuale del progetto (feature già coperte)
Dall'analisi del codice emergono aree già implementate:

- **Connessione Exchange Online + stato sessione** (connect/disconnect/status) tramite worker PowerShell.
- **Mailbox management operativo**: elenco, dettagli, mailbox cancellate, restore, conversione shared/regular, permessi mailbox/recipient, auto-reply, retention policy, creazione mailbox.
- **Distribution list management**: liste statiche/dinamiche, dettagli, membri, modifica membership, creazione DL, settaggi principali.
- **Message Trace**: ricerca e dettaglio tracce messaggi.
- **Mail Flow base**: transport rules, connectors, accepted domains.
- **Licensing via Graph**: lettura licenze utente/disponibili e assegnazione/rimozione.

Queste capability risultano principalmente da `OperationType` e dai metodi presenti in `ExoCommands` / `ExoGroupCommands`.

---

## 2) Gap principali rispetto a una copertura "amministrazione EXO Microsoft-complete"

> Nota: i gap seguenti sono prioritizzati per valore operativo e aderenza alle aree tipiche dell'Exchange Admin Center + Microsoft 365 security/compliance.

### A. Identity & Auth avanzata (alta priorità)
**Presente oggi:** connessione interattiva (`Connect-ExchangeOnline`) e connessione Graph con scope delegati.

**Mancante:**
- **App-only auth con certificato** per automazioni non presidiate.
- **Managed Identity / service principal flow** (scenario enterprise/automation).
- **Scelta esplicita tenant/organization/delegated organization** da UI (CSP/GDAP), oltre alla sola variabile ambiente.
- **Scope Graph minimizzati e configurabili** (principio least privilege).

**Perché è rilevante:** le linee guida Microsoft privilegiano autenticazioni moderne e automation-friendly per attività ripetitive e governance centralizzata.

### B. Sicurezza mail (EOP / Defender) (alta priorità)
**Mancante:**
- Gestione policy **Anti-Spam / Anti-Phish / Anti-Malware / Safe Links / Safe Attachments**.
- Gestione **quarantine** e rilascio messaggi.
- Gestione **DKIM** (enable/rotate), verifica configurazioni SPF/DMARC (almeno check diagnostico).

**Impatto:** area critica per posture security; oggi la sezione mail flow copre solo subset infrastrutturale (rules/connectors/domains).

### C. Compliance & Purview integration (alta priorità)
**Mancante:**
- Ricerca **Unified Audit Log** e auditing amministrativo/operativo.
- Flussi compliance avanzati: hold/eDiscovery case management, DLP/mail compliance policies (in coordinamento Purview).
- Gestione avanzata retention/compliance oltre al set retention policy mailbox.

**Impatto:** senza queste capability l'app non copre molte richieste audit/compliance tipiche enterprise.

### D. Recipient types completi (media-alta priorità)
**Presente oggi:** user/shared mailbox, DL, alcune operazioni su mailbox.

**Mancante:**
- **Mail contacts** e **mail users** lifecycle completo.
- **Resource mailbox** (room/equipment) booking policies complete.
- Gestione estesa **Microsoft 365 Groups** (ownership, settings, policy controls) oltre alla sola capability detection.

### E. Hybrid, migration, coexistence (media priorità)
**Mancante:**
- Gestione **migration batches** (onboarding/offboarding/cutover/staged).
- Configurazioni **hybrid mail flow** e troubleshooting guidato.
- Organization relationships / federation / sharing policies avanzate.

### F. Operatività enterprise (media priorità)
**Mancante:**
- **Bulk operations guidate** (CSV import + preview + dry-run + rollback plan) su più aree.
- **Job scheduler** / execution history per task ricorrenti.
- **Runbook mode** con output firmato/auditabile.
- **Template policy** applicabili per baseline tenant.

### G. Governance/RBAC applicativa (media priorità)
**Mancante:**
- Mappatura e visibilità **RBAC Exchange** (role groups, assignments).
- Policy locali su chi può eseguire quali funzioni nell'app (authorization applicativa).
- Segregazione ruoli Helpdesk/Operator/SecOps nel frontend.

### H. Osservabilità e reportistica (media priorità)
**Mancante:**
- Dashboard KPI evolute (trend, anomalie, SLA, top offenders).
- Export/report pianificati.
- Correlazione errori/cmdlet con suggerimenti remediation più prescrittivi.

---

## 3) Evidenze tecniche rapide
- Le operation implementate coprono mailbox, DL, message trace, transport rules/connectors/domains, licensing e prerequisiti.
- La connessione EXO è costruita su comando base interattivo; Graph usa scope delegati fissi.
- Il detector capability include cmdlet mailbox/group principali ma non aree security/compliance avanzate.

---

## 4) Backlog raccomandato (roadmap)

### Wave 1 (quick win, 2-4 settimane)
1. **Connection profiles** (Interactive / App-only cert) + salvataggio profili tenant.
2. **Least-privilege Graph scopes** configurabili.
3. **Bulk CSV framework** riusabile (preview, validate, execute, rollback report).
4. **Audit trail locale** delle operazioni (who/when/what/result).

### Wave 2 (4-8 settimane)
1. Modulo **Security policies** (EOP/Defender principali).
2. Modulo **Recipient expansion** (contacts, mail users, room/equipment).
3. **RBAC viewer** Exchange (role groups + effective rights snapshot).

### Wave 3 (8+ settimane)
1. **Compliance connector** (audit log search + export).
2. **Migration/Hybrid assistant** con checklist e health checks.
3. Dashboard avanzata con trend e report schedulati.

---

## 5) Raccomandazione sintetica
Il progetto è già solido come **console operativa Exchange Online** per attività day-2 su mailbox, DL, mail trace e una parte del mail flow. Per allinearlo maggiormente a una gestione "secondo specifiche/aspettative Microsoft enterprise", la priorità è estendere:

1. **authentication model** (app-only/automation),
2. **security/compliance surface**,
3. **governance e bulk enterprise operations**.

Questo incremento porterebbe l'app da "tool operativo" a "piattaforma amministrativa completa" per tenant Microsoft 365.
