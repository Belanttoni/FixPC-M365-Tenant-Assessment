# FixPC M365 Tenant Assessment

Script PowerShell **read-only** per la valutazione della sicurezza di tenant Microsoft 365. Raccoglie dati da Microsoft Graph e Exchange Online, calcola un security score e genera report Excel, HTML, Markdown e CSV.

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
| `M365_SecurityAssessment_<tenant>_<date>.html` | **[v2.4]** Report HTML esecutivo in italiano |
| `RapportoEsecutivo_<tenant>_<date>.md` | Report Markdown testuale |
| `assessment.log` | Log completo dell'assessment |
| `CSV/` | Directory con CSV per ogni categoria raccolta |

### CSV generati

- `MFA_Registration.csv` — Stato registrazione MFA utenti
- `Conditional_Access.csv` — Policy Conditional Access
- `SignIn_Logs.csv` — Log di accesso
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

## Architettura

```
src/
  Invoke-M365TenantAssessment.ps1   Script principale (monolitico, auto-contenuto)
Modules/                            Moduli di supporto (estensione futura)
docs/
  Architecture.md                   Documentazione architetturale
```

## Sicurezza e limitazioni

- **Read-only**: lo script non esegue mai operazioni di scrittura sul tenant (nessun `Set-`, `New-`, `Remove-`, `Update-`)
- La raccolta **Risky Users** e **Risk Detections** richiede licenza Azure AD P1/P2 e viene saltata gracefully se non disponibile
- Il timeout per coleta (`-CollectionTimeoutMinutes`) isola ogni raccolta: se una supera il limite viene registrato un `WARN` e l'assessment continua
- Il `Write-Progress` mostra lo stato in tempo reale nelle 7 fasi dell'assessment

## Changelog

Vedere [CHANGELOG.md](CHANGELOG.md)
