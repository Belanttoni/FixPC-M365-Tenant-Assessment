# FixPC M365 Tenant Assessment

Script PowerShell **read-only** per la valutazione della sicurezza di tenant Microsoft 365. Raccoglie dati da Microsoft Graph e Exchange Online, calcola un security score e genera report Excel, HTML, Markdown e CSV — inclusa un'analisi avanzata delle minacce di autenticazione (brute force, password spray, accessi esteri).

## Requisiti

| Componente | Versione minima |
|---|---|
| PowerShell | 7.0+ |
| Microsoft.Graph | 2.0.0 |
| ExchangeOnlineManagement | 3.0.0 |
| ImportExcel | 7.0.0 |

## Installazione moduli

```powershell
Install-Module Microsoft.Graph          -Scope CurrentUser -Force
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
Install-Module ImportExcel              -Scope CurrentUser -Force
```

## Utilizzo

### Assessment completo

```powershell
.\src\Invoke-M365TenantAssessment.ps1 -TenantName "contoso" -LookbackDays 30
```

### Con utente sotto indagine

```powershell
.\src\Invoke-M365TenantAssessment.ps1 -TenantName "contoso" -IncidentUser "john@contoso.com" -LookbackDays 7
```

### Solo Graph (salta Exchange)

```powershell
.\src\Invoke-M365TenantAssessment.ps1 -TenantName "contoso" -SkipExchange
```

### Solo Exchange (salta Graph)

```powershell
.\src\Invoke-M365TenantAssessment.ps1 -TenantName "contoso" -SkipGraph
```

### Verifica sintattica (no connessioni richieste)

```powershell
.\src\Invoke-M365TenantAssessment.ps1 -TenantName dummy -SyntaxOnly
```

## Parametri principali

| Parametro | Default | Descrizione |
|---|---|---|
| `-TenantName` | *(obbligatorio)* | Nome o dominio del tenant M365 |
| `-OutputPath` | `.\M365Assessment_<timestamp>` | Directory di output |
| `-LookbackDays` | `7` | Giorni di lookback per sign-in logs (1-90) |
| `-CollectionTimeoutMinutes` | `5` | Timeout per ogni singola raccolta dati (1-60) |
| `-IncidentUser` | — | UPN utente per indagine specifica |
| `-IncludeRawData` | `$false` | Includi foglio dati grezzi nel workbook Excel |
| `-SkipGraph` | `$false` | Salta raccolta Microsoft Graph |
| `-SkipExchange` | `$false` | Salta raccolta Exchange Online |
| `-SyntaxOnly` | `$false` | Solo verifica sintattica, nessuna connessione |

## Output generati

Tutti i file vengono salvati nella directory di output (`-OutputPath`):

| File | Descrizione |
|---|---|
| `M365_SecurityAssessment_<tenant>_<date>.xlsx` | Workbook Excel multi-foglio con dati completi |
| `M365_SecurityAssessment_<tenant>_<date>.html` | **[v2.4+]** Report HTML esecutivo in italiano |
| `RapportoEsecutivo_<tenant>_<date>.md` | Report Markdown testuale |
| `assessment.log` | Log completo dell'assessment |
| `CSV/` | Directory con CSV per ogni categoria raccolta |

### Fogli Excel

Il workbook contiene fino a 14 fogli (13 + Raw Data opzionale):

| Foglio | Contenuto |
|---|---|
| Sintesi Esecutiva | Score complessivo, info tenant, riepilogo rischi |
| Score Sicurezza | Punteggio dettagliato per controllo |
| Azioni Prioritarie | Piano d'azione con priorità |
| MFA Coverage | Stato MFA per utente |
| Conditional Access | Policy CA configurate |
| Legacy Authentication | Eventi autenticazione legacy |
| SMTP AUTH POP IMAP | Configurazione protocolli a rischio |
| Global Administrators | Ruoli privilegiati e membri |
| OAuth Grants | Grant OAuth con livello di rischio |
| Sign-in Anomalies | Anomalie sign-in e accessi a rischio |
| Mailbox Rules Forwarding | Regole inbox e forward sospetti |
| Risk Detections | Rilevamenti Identity Protection (P1/P2) |
| **Authentication Threats** | **[v2.5]** Analisi minacce: accessi esteri, brute force, password spray, SAF, CA notApplied |
| Raw Data | Statistiche raccolta (solo con `-IncludeRawData`) |

### CSV generati

**CSV raccolta dati (esistenti):**
- `MFA_Registration.csv` — Stato registrazione MFA utenti
- `Conditional_Access.csv` — Policy Conditional Access
- `SignIn_Logs.csv` — Log di accesso (con campi `Country` e `City` separati da v2.5)
- `Legacy_Auth.csv` — Autenticazione legacy rilevata
- `Transport_Config.csv` — Configurazione SMTP/TLS tenant
- `CAS_Mailboxes.csv` — Impostazioni CAS mailbox (POP/IMAP/SMTP)
- `Mailboxes.csv` — Elenco mailbox con forwarding
- `Directory_Roles.csv` — Ruoli privilegiati e membri
- `OAuth_Grants.csv` — Grant OAuth (sempre generato, anche vuoto)
- `Risky_Users.csv` — Utenti a rischio (richiede licenza P2)
- `Risk_Detections.csv` — Rilevamenti rischi (richiede licenza P1/P2)
- `Inbox_Rules.csv` — Regole inbox potenzialmente rischiose
- `Forwarding.csv` — Configurazioni forwarding mailbox

**CSV analisi minacce autenticazione (nuovi in v2.5, sempre generati anche vuoti):**
- `Foreign_SignIns.csv` — Tutti i login con paese diverso da IT (Success + Failure)
- `Failed_Login_Analysis.csv` — Fallimenti raggruppati per UPN/IP/Paese/App con conteggio
- `Brute_Force_Candidates.csv` — Utenti con ≥10 fallimenti in 15 minuti (severità High)
- `Password_Spray_Candidates.csv` — IP con ≥5 utenti distinti tentati (severità High)
- `Successful_After_Failures.csv` — Login riuscito dopo ≥2 fallimenti in 60 min (Critical se paese estero)
- `Conditional_Access_NotApplied.csv` — Eventi con CA non applicato, successi evidenziati

## Architettura

```
src/
  Invoke-M365TenantAssessment.ps1   Script principale (monolitico, auto-contenuto)
Modules/                            Moduli di supporto (estensione futura)
docs/
  Architecture.md                   Documentazione architetturale
```

## Analisi minacce di autenticazione (v2.5)

La funzione `Get-AuthenticationThreatAnalysis` analizza i sign-in log già raccolti (senza chiamate Graph aggiuntive) e produce 6 dataset:

| Dataset | Logica di rilevamento | Severità |
|---|---|---|
| Foreign SignIns | `CountryOrRegion != IT` | — |
| Failed Login Analysis | Raggruppamento per UPN+IP+Paese+App | — |
| Brute Force Candidates | ≥10 fallimenti stesso utente in 15 min | **High** |
| Password Spray Candidates | Stesso IP vs ≥5 utenti distinti | **High** |
| Successful After Failures | ≥2 fallimenti + successo in 60 min | **High** / **Critical** (paese estero) |
| CA NotApplied | `ConditionalAccessStatus = notApplied` | — |

Se non ci sono sign-in log (es. `-SkipGraph` attivo o timeout), tutti i dataset sono vuoti e i CSV vengono comunque generati con un placeholder.

## Sicurezza e limitazioni

- **Read-only**: lo script non esegue mai operazioni di scrittura sul tenant (nessun `Set-`, `New-`, `Remove-`, `Update-`)
- La raccolta **Risky Users** e **Risk Detections** richiede licenza Azure AD P1/P2 e viene saltata gracefully se non disponibile
- Il timeout per coleta (`-CollectionTimeoutMinutes`) isola ogni raccolta: se una supera il limite viene registrato un `WARN` e l'assessment continua
- Il `Write-Progress` mostra lo stato in tempo reale nelle 7 fasi dell'assessment
- L'analisi minacce opera **offline** sui dati già raccolti: nessuna chiamata Graph aggiuntiva, nessun impatto sul timeout

## Changelog

Vedere [CHANGELOG.md](CHANGELOG.md)
