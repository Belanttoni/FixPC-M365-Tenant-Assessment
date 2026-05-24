# Changelog

Tutte le modifiche rilevanti al progetto FixPC M365 Tenant Assessment.

Il formato segue [Keep a Changelog](https://keepachangelog.com/it/1.0.0/).

---

## [2.5] - 2026-05-24

### Aggiunto

- **`Get-AuthenticationThreatAnalysis`**: nuova funzione che analizza i sign-in log già raccolti da `Get-MgAuditLogSignIn` senza ulteriori chiamate a Graph. Produce 6 dataset distinti:
  - **`Foreign_SignIns`**: tutti i login con `CountryOrRegion != IT`, con campo `Result` (Success/Failure), separati per analisi.
  - **`Failed_Login_Analysis`**: tutte le autenticazioni fallite raggruppate per UPN, IP, Paese, App con conteggio totale fallimenti e campione di errori.
  - **`Brute_Force_Candidates`**: utenti con ≥10 fallimenti in una finestra di 15 minuti — severità **High**.
  - **`Password_Spray_Candidates`**: IP che tentano ≥5 utenti distinti nel periodo analizzato — severità **High**.
  - **`Successful_After_Failures`**: utente con ≥2 fallimenti seguiti da login riuscito in 60 minuti — severità **Critical** se paese estero.
  - **`Conditional_Access_NotApplied`**: eventi con `ConditionalAccessStatus = notApplied`, con login riusciti evidenziati.
- **Campi `Country` e `City`** aggiunti all'output di `Get-SignInLogsData` (prima erano uniti nel campo `Location`). Non modifica la raccolta Graph esistente.
- **Foglio Excel `Authentication Threats`** (ws13): riepilogo KPI minacce + 6 sezioni dettagliate con colorazione per severità (Critical = rosso, High = arancio, Medium = giallo, OK = verde). Il vecchio foglio Raw Data diventa ws14.
- **Sezione HTML `Minacce di Autenticazione`**: 6 KPI card colorati dinamicamente (verde/giallo/arancio/rosso) nel report HTML esecutivo — mostra logins fuori Italia, gruppi fallimenti, IP sospetti, brute force, successi dopo fallimenti, CA notApplied.
- **6 nuovi CSV** nella directory `CSV/`: `Foreign_SignIns.csv`, `Failed_Login_Analysis.csv`, `Brute_Force_Candidates.csv`, `Password_Spray_Candidates.csv`, `Successful_After_Failures.csv`, `Conditional_Access_NotApplied.csv`. Tutti sempre generati (anche vuoti con placeholder) per garantire la presenza nel workflow di post-processing.
- **Helper interno `ConvertTo-SafeDateTime`** nella funzione di analisi per gestione robusta di `DateTime`/`DateTimeOffset` provenienti da Graph.

### Modificato

- **`Export-DataToCSV`**: aggiunto parametro `-ThreatData` e 6 nuove voci nel `$csvMap`. Tutti i 6 CSV threat aggiunti a `$alwaysExport` con messaggi placeholder specifici per tipo.
- **`New-ExcelWorkbook`**: aggiunto parametro `-ThreatData` e nuovo foglio `Authentication Threats` (ws13). Il foglio `Raw Data` opzionale rinominato ws14 internamente.
- **`New-HtmlReport`**: aggiunto parametro `-ThreatData` e nuova sezione `Minacce di Autenticazione` con 6 KPI card a colori dinamici.
- **`Main`**: aggiunta chiamata a `Get-AuthenticationThreatAnalysis` nella Fase 3 (dopo `Get-LegacyAuthData`). Il risultato `$threatData` passato a `Export-DataToCSV`, `New-ExcelWorkbook` e `New-HtmlReport`.
- **Versione**: aggiornata da `2.4` a `2.5`.

### Invariato

- Script 100% **read-only**: nessuna chiamata a `Set-`, `New-`, `Remove-`, `Update-` su servizi cloud.
- `-SkipGraph`, `-SkipExchange` e `-SyntaxOnly` continuano a funzionare esattamente come prima. Se `-SkipGraph` è attivo, `$threatData` sarà `$null` e tutti i 6 CSV vengono generati vuoti senza errori.
- Se non ci sono sign-in log (timeout, errore, zero eventi), `Get-AuthenticationThreatAnalysis` restituisce dataset vuoti e il processo continua senza interruzione.
- Tutti i file di output pre-esistenti (Excel, Markdown, CSV legacy) mantengono struttura invariata.
- Requisiti moduli invariati: `Microsoft.Graph >= 2.0`, `ExchangeOnlineManagement >= 3.0`, `ImportExcel >= 7.0`.

---

## [2.4] - 2025-05-24

### Aggiunto

- **HTML Report esecutivo** (`New-HtmlReport`): genera `M365_SecurityAssessment_<tenant>_<date>.html` con layout visuale in italiano, score gauge, KPI grid, tabella rischi con badge colorati, tabella azioni prioritarie e disclaimer. Nessun dato sensibile esteso incluso.
- **Parametro `-CollectionTimeoutMinutes`** (default: 5): imposta il timeout in minuti per ogni singola raccolta dati. Ogni coleta viene eseguita in un `Start-ThreadJob` isolato con `Wait-Job -Timeout`; se supera il limite viene registrato un `[TIMEOUT] WARN` e l'assessment continua senza interruzione.
- **Progress bars** (`Write-Progress`) per tutte le fasi principali: verifica moduli, connessione Graph, connessione Exchange, raccolta Graph, raccolta Exchange, calcolo score, export CSV, generazione Excel, generazione report Markdown, generazione HTML.
- **Progress per mailbox** in `Get-MailboxRulesData`: ogni mailbox mostra avanzamento percentuale durante l'analisi delle inbox rules.
- **`OAuth_Grants.csv` sempre generato**: anche in assenza di dati (timeout, errore o zero grant), il file viene creato con una riga placeholder per garantire la presenza del file nel workflow di post-processing.
- **Costante `$Script:CollectionTimeoutMinutes`**: inizializzata dal parametro `-CollectionTimeoutMinutes` nella sezione costanti.

### Modificato

- **`Invoke-SafeCollection`**: rewritten con timeout engine. Tenta l'esecuzione in `Start-ThreadJob` (PS7, isolation reale); se il job fallisce per mancato contesto sessione (EXO/Graph), fallback automatico all'esecuzione inline con rilevamento tempo. Il WARN viene registrato in entrambi i casi di timeout/lentezza.
- **`Get-OAuthGrantsData`**: handling completamente difensivo. Ogni grant viene elaborato in un `try/catch` interno; le proprietà `ExpiryTime`, `PublisherName`, `ConsentType`, `Scope` vengono lette con `PSObject.Properties.Name -contains` e fallback su `AdditionalProperties`. Grant con struttura anomala vengono saltati con `Write-Warning` senza interrompere la raccolta.
- **`Export-DataToCSV`**: aggiunta logica `$alwaysExport` per garantire la presenza di file CSV critici anche vuoti. Aggiunto `Write-Progress` per tracciare l'avanzamento dell'export.
- **`$ProgressPreference`**: cambiato da `SilentlyContinue` a `Continue` per abilitare la visualizzazione delle progress bar.
- **`Main`**: suddivisa in 7 fasi esplicite con `Write-Progress`. Aggiunta chiamata a `New-HtmlReport` e `HtmlReport` nell'oggetto di output finale.
- **Versione**: aggiornata da `2.3` a `2.4` in `$Script:AssessmentVersion`, nel `.NOTES` dell'help e in tutti i riferimenti nei report.
- **README.md**: aggiornato con parametro `-CollectionTimeoutMinutes`, tabella output aggiornata (HTML report), sezione CSV aggiornata (`OAuth_Grants.csv` sempre generato).

### Invariato

- Script 100% **read-only**: nessun uso di `Set-`, `New-`, `Remove-`, `Update-` su servizi cloud.
- `-SkipGraph`, `-SkipExchange` e `-SyntaxOnly` continuano a funzionare esattamente come prima.
- Tutti i file di output pre-esistenti (Excel, Markdown, CSV) sono mantenuti senza modifiche alla struttura.
- Requisiti moduli invariati: `Microsoft.Graph >= 2.0`, `ExchangeOnlineManagement >= 3.0`, `ImportExcel >= 7.0`.

---

## [2.3] - 2025-05-01

### Aggiunto

- Script monolitico `Invoke-M365TenantAssessment.ps1` con raccolta da Graph e Exchange Online.
- Workbook Excel multi-foglio (12-13 fogli) con score, azioni, MFA, CA, OAuth, Sign-in, ecc.
- Report Markdown esecutivo (`RapportoEsecutivo_<tenant>_<date>.md`).
- Export CSV per tutte le categorie di dati raccolti.
- Parametro `-SyntaxOnly` per verifica sintattica offline.
- Parametri `-SkipGraph` e `-SkipExchange` per raccolta parziale.
- Gestione graceful per licenze P1/P2 mancanti (Risky Users, Risk Detections).
- `Invoke-SafeCollection` come wrapper error-safe per ogni raccolta.
- Security score calcolato su 7 controlli (MFA, SMTP, CA, Legacy Auth, Admin, POP/IMAP, OAuth).
