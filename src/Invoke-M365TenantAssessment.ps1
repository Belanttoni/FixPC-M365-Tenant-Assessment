#Requires -Version 7.0
<#
.SYNOPSIS
    Invoke-M365TenantAssessment.ps1 - Read-Only Security Assessment for Microsoft 365 Tenants

.DESCRIPTION
    Esegue un assessment di sicurezza read-only su un tenant Microsoft 365.
    Raccoglie dati da Microsoft Graph e Exchange Online, esporta CSV e genera
    un workbook Excel esecutivo in italiano con score di sicurezza, rischi e piano d'azione.

.PARAMETER TenantName
    Nome del tenant Microsoft 365 (es. contoso o contoso.onmicrosoft.com)

.PARAMETER OutputPath
    Percorso della directory di output. Default: .\M365Assessment_<timestamp>

.PARAMETER LookbackDays
    Numero di giorni da analizzare per i log di accesso. Default: 30

.PARAMETER IncidentUser
    UPN dell'utente oggetto di indagine (opzionale)

.PARAMETER IncludeRawData
    Switch per includere i dati grezzi nel workbook Excel

.PARAMETER SkipExchange
    Switch per saltare la raccolta dati Exchange Online (utile per test parziali)

.PARAMETER SkipGraph
    Switch per saltare la raccolta dati Microsoft Graph (utile per test parziali)

.EXAMPLE
    .\Invoke-M365TenantAssessment.ps1 -TenantName "contoso" -OutputPath "C:\Assessments\Contoso" -LookbackDays 30

.EXAMPLE
    .\Invoke-M365TenantAssessment.ps1 -TenantName "contoso" -IncidentUser "john.doe@contoso.com" -IncludeRawData

.PARAMETER CollectionTimeoutMinutes
    Timeout in minuti per ogni singola coleta di dati. Default: 5.
    Se una coleta supera il limite, viene registrato un WARN e l'assessment continua.

.PARAMETER SyntaxOnly
    Switch per eseguire solo la verifica sintattica del file e uscire immediatamente.
    Non richiede connessioni o moduli extra oltre al parser PowerShell nativo.
    Esempio: .\Invoke-M365TenantAssessment.ps1 -TenantName dummy -SyntaxOnly

.EXAMPLE
    .\Invoke-M365TenantAssessment.ps1 -TenantName "contoso" -SkipExchange -WhatIf

.EXAMPLE
    .\Invoke-M365TenantAssessment.ps1 -TenantName dummy -SyntaxOnly

.NOTES
    Autore: Security Assessment Script
    Versione: 2.4
    Requisiti: PowerShell 7+, Microsoft.Graph >= 2.0, ExchangeOnlineManagement >= 3.0,
               ImportExcel >= 7.0 (obbligatorio: carica EPPlus/OfficeOpenXml usato internamente)
    IMPORTANTE: Script read-only - non apporta modifiche al tenant.
    NOTA EXCEL: Il modulo ImportExcel e obbligatorio anche se non si usa Export-Excel
                direttamente, perche il suo import carica l'assembly EPPlus
                (OfficeOpenXml) usato dalle funzioni di formattazione del workbook.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Nome del tenant M365 (es. contoso)")]
    [ValidateNotNullOrEmpty()]
    [string]$TenantName,

    [Parameter(Mandatory = $false, HelpMessage = "Percorso output")]
    [string]$OutputPath = "",

    [Parameter(Mandatory = $false, HelpMessage = "Giorni lookback per sign-in logs")]
    [ValidateRange(1, 90)]
    [int]$LookbackDays = 7,

    [Parameter(Mandatory = $false, HelpMessage = "UPN utente per indagine specifica")]
    [string]$IncidentUser = "",

    [Parameter(Mandatory = $false, HelpMessage = "Includi dati grezzi nel workbook Excel")]
    [switch]$IncludeRawData,

    [Parameter(Mandatory = $false, HelpMessage = "Salta raccolta Exchange Online (test parziale)")]
    [switch]$SkipExchange,

    [Parameter(Mandatory = $false, HelpMessage = "Salta raccolta Microsoft Graph (test parziale)")]
    [switch]$SkipGraph,

    [Parameter(Mandatory = $false, HelpMessage = "Timeout in minuti per ogni singola raccolta dati (default 5)")]
    [ValidateRange(1, 60)]
    [int]$CollectionTimeoutMinutes = 5,

    [Parameter(Mandatory = $false, HelpMessage = "Esegue solo verifica sintattica del file e termina (non richiede connessioni)")]
    [switch]$SyntaxOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"
$ProgressPreference    = "Continue"

# Imposta OutputPath di default se non specificato (non usabile nel param default con Get-Date)
if ([string]::IsNullOrEmpty($OutputPath)) {
    $OutputPath = ".\M365Assessment_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
}

#region ─── COSTANTI E CONFIGURAZIONE ──────────────────────────────────────────

$Script:AssessmentVersion        = "2.4"
$Script:StartTime                = Get-Date
$Script:LogFile                  = $null
$Script:OutputDir                = $null
$Script:CollectionTimeoutMinutes = $CollectionTimeoutMinutes

# Soglie per il calcolo del security score
$Script:ScoreThresholds = @{
    SmtpAuthDisabled = 15
    MfaHighCoverage  = 20
    MfaMedCoverage   = 10
    CaMfaEnabled     = 20
    BlockLegacyAuth  = 15
    FewGlobalAdmins  = 10
    PoPImapDisabled  = 10
    LowOauthRisk     = 10
}

# Scope Graph richiesti
$Script:GraphScopes = @(
    "AuditLog.Read.All"
    "Directory.Read.All"
    "Policy.Read.All"
    "User.Read.All"
    "UserAuthenticationMethod.Read.All"
    "Reports.Read.All"
    "Application.Read.All"
    "DelegatedPermissionGrant.Read.All"
    "IdentityRiskyUser.Read.All"
    "IdentityRiskEvent.Read.All"
)

# Scope OAuth considerati ad alto rischio
$Script:HighRiskOAuthScopes = @(
    "Mail.ReadWrite", "Mail.Send", "MailboxSettings.ReadWrite",
    "Files.ReadWrite.All", "Sites.FullControl.All", "Directory.ReadWrite.All",
    "RoleManagement.ReadWrite.Directory", "AppRoleAssignment.ReadWrite.All",
    "Application.ReadWrite.All", "User.ReadWrite.All", "GroupMember.ReadWrite.All"
)

# Colori Excel (hex senza #)
$Script:Colors = @{
    HeaderBg   = "1F4E79"
    HeaderFg   = "FFFFFF"
    CriticalBg = "C00000"
    CriticalFg = "FFFFFF"
    HighBg     = "FF0000"
    HighFg     = "FFFFFF"
    MediumBg   = "FFC000"
    MediumFg   = "000000"
    LowBg      = "FFFF00"
    LowFg      = "000000"
    OkBg       = "70AD47"
    OkFg       = "FFFFFF"
    InfoBg     = "BDD7EE"
    InfoFg     = "000000"
    SectionBg  = "2E75B6"
    SectionFg  = "FFFFFF"
    AltRowBg   = "DEEAF1"
    WhiteBg    = "FFFFFF"
    TitleBg    = "1F4E79"
    TitleFg    = "FFFFFF"
}

#endregion

#region ─── TEST SINTASSI ───────────────────────────────────────────────────────

function Test-ScriptSyntax {
    <#
    .SYNOPSIS
        Verifica la sintassi PowerShell del file indicato usando il parser nativo.
    .PARAMETER Path
        Percorso del file .ps1 da verificare.
    .OUTPUTS
        PSCustomObject con IsValid, ErrorCount, Errors.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        [string]$Path
    )

    $parseErrors = $null
    $null = [System.Management.Automation.Language.Parser]::ParseFile(
        (Resolve-Path $Path).Path,
        [ref]$null,
        [ref]$parseErrors
    )

    $result = [PSCustomObject]@{
        IsValid    = ($parseErrors.Count -eq 0)
        ErrorCount = $parseErrors.Count
        Errors     = $parseErrors
        Path       = $Path
    }

    if ($result.IsValid) {
        Write-Host "[Test-ScriptSyntax] PASS - Nessun errore di sintassi in: $Path" -ForegroundColor Green
    } else {
        Write-Host "[Test-ScriptSyntax] FAIL - $($result.ErrorCount) errore/i in: $Path" -ForegroundColor Red
        foreach ($err in $parseErrors) {
            Write-Host "  Riga $($err.Extent.StartLineNumber): $($err.Message)" -ForegroundColor Yellow
        }
    }
    return $result
}

#endregion

#region ─── FUNZIONI DI LOGGING ────────────────────────────────────────────────

function Initialize-Logging {
    param([string]$LogPath)
    $Script:LogFile = $LogPath
    $header = "================================================================================" + [Environment]::NewLine +
              "  M365 TENANT SECURITY ASSESSMENT v$($Script:AssessmentVersion)" + [Environment]::NewLine +
              "  Tenant: $TenantName" + [Environment]::NewLine +
              "  Avviato: $($Script:StartTime.ToString('yyyy-MM-dd HH:mm:ss'))" + [Environment]::NewLine +
              "  Operatore: $($env:USERNAME)@$($env:COMPUTERNAME)" + [Environment]::NewLine +
              "================================================================================"
    $header | Set-Content -Path $LogPath -Encoding UTF8
    Write-Host $header -ForegroundColor Cyan
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS", "SECTION")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine   = "[$timestamp] [$Level] $Message"

    if ($Script:LogFile) {
        Add-Content -Path $Script:LogFile -Value $logLine -Encoding UTF8
    }

    $color = switch ($Level) {
        "INFO"    { "White"  }
        "WARN"    { "Yellow" }
        "ERROR"   { "Red"    }
        "SUCCESS" { "Green"  }
        "SECTION" { "Cyan"   }
        default   { "White"  }
    }
    Write-Host $logLine -ForegroundColor $color
}

function Write-Section {
    param([string]$Title)
    $line = "-" * 70
    Write-Log "" "INFO"
    Write-Log $line "SECTION"
    Write-Log "  $Title" "SECTION"
    Write-Log $line "SECTION"
}

#endregion

#region ─── VERIFICA E CARICAMENTO MODULI ──────────────────────────────────────

function Test-AndImportModules {
    Write-Section "VERIFICA MODULI POWERSHELL"

    $requiredModules = @(
        @{ Name = "Microsoft.Graph";          MinVersion = "2.0.0" }
        @{ Name = "ExchangeOnlineManagement"; MinVersion = "3.0.0" }
        @{ Name = "ImportExcel";              MinVersion = "7.0.0" }
    )

    $allOk    = $true
    $modIndex = 0
    foreach ($mod in $requiredModules) {
        $modIndex++
        Write-Progress -Activity "Verifica moduli PowerShell" `
            -Status "Controllo: $($mod.Name)" `
            -PercentComplete ([math]::Round(($modIndex / $requiredModules.Count) * 100))

        # Salta moduli non necessari in base ai parametri
        if ($SkipGraph   -and $mod.Name -eq "Microsoft.Graph")          { continue }
        if ($SkipExchange -and $mod.Name -eq "ExchangeOnlineManagement") { continue }

        $installed = Get-Module -ListAvailable -Name $mod.Name |
            Sort-Object Version -Descending | Select-Object -First 1

        if (-not $installed) {
            Write-Log "Modulo mancante: $($mod.Name) (minimo v$($mod.MinVersion))" "ERROR"
            Write-Log "Installare con: Install-Module -Name $($mod.Name) -Scope CurrentUser -Force" "WARN"
            $allOk = $false
        } elseif ([Version]$installed.Version -lt [Version]$mod.MinVersion) {
            Write-Log "Modulo $($mod.Name) v$($installed.Version) troppo vecchio (minimo v$($mod.MinVersion))" "ERROR"
            Write-Log "Aggiornare con: Update-Module -Name $($mod.Name)" "WARN"
            $allOk = $false
        } else {
            Write-Log "OK: $($mod.Name) v$($installed.Version)" "SUCCESS"
            try {
                Import-Module $mod.Name -ErrorAction Stop
            } catch {
                Write-Log "Errore caricamento $($mod.Name): $($_.Exception.Message)" "ERROR"
                $allOk = $false
            }
        }
    }
    Write-Progress -Activity "Verifica moduli PowerShell" -Completed

    if (-not $allOk) {
        throw "Uno o piu moduli richiesti non sono disponibili. Installare i moduli mancanti e riprovare."
    }
    Write-Log "Tutti i moduli verificati e caricati." "SUCCESS"
}

#endregion

#region ─── CONNESSIONI ────────────────────────────────────────────────────────

function Connect-ToMicrosoftGraph {
    Write-Section "CONNESSIONE MICROSOFT GRAPH"
    Write-Log "Scope richiesti: $($Script:GraphScopes -join ', ')" "INFO"

    Write-Progress -Activity "Connessione Microsoft Graph" -Status "Verifica contesto esistente..." -PercentComplete 10
    try {
        $ctx = Get-MgContext -ErrorAction SilentlyContinue
        if ($ctx -and $ctx.TenantId) {
            Write-Log "Gia connesso al Graph come: $($ctx.Account)" "INFO"
            $missingScopes = $Script:GraphScopes | Where-Object { $_ -notin $ctx.Scopes }
            if ($missingScopes.Count -gt 0) {
                Write-Log "Scope mancanti, ri-connessione necessaria: $($missingScopes -join ', ')" "WARN"
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            } else {
                Write-Log "Scope verificati. Connessione valida." "SUCCESS"
                Write-Progress -Activity "Connessione Microsoft Graph" -Completed
                return
            }
        }

        Write-Progress -Activity "Connessione Microsoft Graph" -Status "Autenticazione in corso..." -PercentComplete 50
        Connect-MgGraph -Scopes $Script:GraphScopes -TenantId $TenantName -NoWelcome -ErrorAction Stop
        $ctx = Get-MgContext
        Write-Log "Connesso al Graph come: $($ctx.Account)" "SUCCESS"
        Write-Log "Tenant ID: $($ctx.TenantId)" "INFO"
        Write-Progress -Activity "Connessione Microsoft Graph" -Completed
    } catch {
        Write-Progress -Activity "Connessione Microsoft Graph" -Completed
        throw "Connessione Microsoft Graph fallita: $($_.Exception.Message)"
    }
}

function Connect-ToExchangeOnline {
    Write-Section "CONNESSIONE EXCHANGE ONLINE"

    Write-Progress -Activity "Connessione Exchange Online" -Status "Verifica sessione esistente..." -PercentComplete 20
    try {
        $existingSession = Get-PSSession | Where-Object {
            $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened"
        }
        if ($existingSession) {
            Write-Log "Gia connesso a Exchange Online." "INFO"
            Write-Progress -Activity "Connessione Exchange Online" -Completed
            return
        }

        $tenantDomain = if ($TenantName -like "*.onmicrosoft.com" -or $TenantName -like "*.*") {
            $TenantName
        } else {
            "$TenantName.onmicrosoft.com"
        }

        Write-Progress -Activity "Connessione Exchange Online" -Status "Autenticazione per $tenantDomain..." -PercentComplete 60
        Connect-ExchangeOnline -Organization $tenantDomain -ShowBanner:$false -ErrorAction Stop
        Write-Log "Connesso a Exchange Online per: $tenantDomain" "SUCCESS"
        Write-Progress -Activity "Connessione Exchange Online" -Completed
    } catch {
        Write-Progress -Activity "Connessione Exchange Online" -Completed
        throw "Connessione Exchange Online fallita: $($_.Exception.Message)"
    }
}

#endregion

#region ─── RACCOLTA DATI ──────────────────────────────────────────────────────

function Invoke-SafeCollection {
    <#
    .SYNOPSIS
        Esegue una raccolta dati con timeout isolato per coleta.
    .DESCRIPTION
        Tenta l'esecuzione in un ThreadJob (PS7) per isolation e timeout reale.
        Se il ThreadJob non riesce (contesto sessione non condiviso), esegue inline.
        In entrambi i casi: registra WARN se timeout, non interrompe l'assessment.
    #>
    param(
        [string]$CollectionName,
        [scriptblock]$ScriptBlock,
        [int]$TimeoutSec = ($Script:CollectionTimeoutMinutes * 60)
    )

    Write-Log "Raccolta: $CollectionName (timeout: $([math]::Round($TimeoutSec/60,1)) min)..." "INFO"
    $sw  = [System.Diagnostics.Stopwatch]::StartNew()
    $job = $null

    try {
        # Tentativo con ThreadJob per timeout reale
        $job      = Start-ThreadJob -ScriptBlock $ScriptBlock -ErrorAction Stop
        $completed = $job | Wait-Job -Timeout $TimeoutSec

        if ($null -eq $completed) {
            # TIMEOUT: interrompi il job e segnala WARN
            $job | Stop-Job -ErrorAction SilentlyContinue
            $job | Remove-Job -Force -ErrorAction SilentlyContinue
            $sw.Stop()
            Write-Log "[TIMEOUT] '$CollectionName' ha superato $([math]::Round($TimeoutSec/60,1)) min - dati non raccolti, assessment continua." "WARN"
            return $null
        }

        if ($job.State -eq 'Failed') {
            $jobErr = ($job.ChildJobs | ForEach-Object { $_.Error } | Select-Object -First 1)
            $job | Remove-Job -Force -ErrorAction SilentlyContinue
            $errMsg = if ($jobErr) { $jobErr.Exception.Message } else { "Errore sconosciuto nel thread job" }
            throw $errMsg
        }

        $result = $job | Receive-Job -ErrorAction SilentlyContinue
        $job | Remove-Job -Force -ErrorAction SilentlyContinue
        $sw.Stop()

        $count = if ($result -is [Array]) { $result.Count } elseif ($null -eq $result) { 0 } else { 1 }
        Write-Log "Completato: $CollectionName - $count elementi ($([math]::Round($sw.Elapsed.TotalSeconds))s)." "SUCCESS"
        return $result

    } catch {
        # Fallback inline: il ThreadJob potrebbe non condividere il contesto EXO/Graph
        if ($job) { $job | Remove-Job -Force -ErrorAction SilentlyContinue }
        $jobErrMsg = $_.Exception.Message

        Write-Log "ThreadJob non disponibile per '$CollectionName' ($jobErrMsg) - esecuzione inline..." "WARN"
        $sw.Restart()

        try {
            $result = & $ScriptBlock
            $sw.Stop()

            if ($sw.Elapsed.TotalSeconds -gt $TimeoutSec) {
                Write-Log "[LENTO] '$CollectionName' ha impiegato $([math]::Round($sw.Elapsed.TotalSeconds))s (soglia: $([math]::Round($TimeoutSec/60,1)) min)." "WARN"
            }
            $count = if ($result -is [Array]) { $result.Count } elseif ($null -eq $result) { 0 } else { 1 }
            Write-Log "Completato (inline): $CollectionName - $count elementi ($([math]::Round($sw.Elapsed.TotalSeconds))s)." "SUCCESS"
            return $result

        } catch {
            $sw.Stop()
            Write-Log "ERRORE in '$CollectionName': $($_.Exception.Message)" "ERROR"
            return $null
        }
    }
}

function Get-MFARegistrationData {
    return Invoke-SafeCollection -CollectionName "MFA Registration Report" -ScriptBlock {
        $raw = Get-MgReportAuthenticationMethodUserRegistrationDetail -All -ErrorAction Stop
        $raw | Select-Object `
            @{N = "UPN";                 E = { $_.UserPrincipalName }},
            @{N = "DisplayName";         E = { $_.UserDisplayName }},
            @{N = "IsAdmin";             E = { $_.IsAdmin }},
            @{N = "IsMfaRegistered";     E = { $_.IsMfaRegistered }},
            @{N = "IsMfaCapable";        E = { $_.IsMfaCapable }},
            @{N = "IsPasswordlessCapable"; E = { $_.IsPasswordlessCapable }},
            @{N = "IsSsprRegistered";    E = { $_.IsSsprRegistered }},
            @{N = "DefaultMfaMethod";    E = { $_.DefaultMfaMethod }},
            @{N = "MethodsRegistered";   E = { ($_.MethodsRegistered -join "; ") }}
    }
}

function Get-ConditionalAccessData {
    return Invoke-SafeCollection -CollectionName "Conditional Access Policies" -ScriptBlock {
        $policies = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop
        $policies | Select-Object `
            @{N = "PolicyId";           E = { $_.Id }},
            @{N = "DisplayName";        E = { $_.DisplayName }},
            @{N = "State";              E = { $_.State }},
            @{N = "IncludeUsers";       E = { ($_.Conditions.Users.IncludeUsers -join "; ") }},
            @{N = "ExcludeUsers";       E = { ($_.Conditions.Users.ExcludeUsers -join "; ") }},
            @{N = "IncludeGroups";      E = { ($_.Conditions.Users.IncludeGroups -join "; ") }},
            @{N = "IncludeApplications"; E = { ($_.Conditions.Applications.IncludeApplications -join "; ") }},
            @{N = "IncludePlatforms";   E = { ($_.Conditions.Platforms.IncludePlatforms -join "; ") }},
            @{N = "ClientAppTypes";     E = { ($_.Conditions.ClientAppTypes -join "; ") }},
            @{N = "GrantControls";      E = { ($_.GrantControls.BuiltInControls -join "; ") }},
            @{N = "GrantOperator";      E = { $_.GrantControls.Operator }},
            @{N = "SessionSignInFreq";  E = { $_.SessionControls.SignInFrequency.Value }},
            @{N = "SessionCAE";         E = { $_.SessionControls.ContinuousAccessEvaluation.Mode }},
            @{N = "CreatedDateTime";    E = { $_.CreatedDateTime }},
            @{N = "ModifiedDateTime";   E = { $_.ModifiedDateTime }}
    }
}

function Get-SignInLogsData {
    param([int]$Days)
    return Invoke-SafeCollection -CollectionName "Sign-in Logs ($Days giorni)" -ScriptBlock {
        $startDate = (Get-Date).AddDays(-$Days).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $filter    = "createdDateTime ge $startDate"

        if (-not [string]::IsNullOrEmpty($IncidentUser)) {
            $filter += " and userPrincipalName eq '$IncidentUser'"
            Write-Log "Filtro utente incidente: $IncidentUser" "INFO"
        }

        $rawSignIns = Get-MgAuditLogSignIn -Filter $filter -Top 999 -All -ErrorAction Stop

        $rawSignIns | Select-Object `
            @{N = "UPN";                    E = { $_.UserPrincipalName }},
            @{N = "DisplayName";            E = { $_.UserDisplayName }},
            @{N = "AppDisplayName";         E = { $_.AppDisplayName }},
            @{N = "ClientAppUsed";          E = { $_.ClientAppUsed }},
            @{N = "IPAddress";              E = { $_.IpAddress }},
            @{N = "Location";               E = { if ($_.Location) { "$($_.Location.City), $($_.Location.CountryOrRegion)" } else { "N/A" } }},
            @{N = "Status";                 E = { $_.Status.ErrorCode }},
            @{N = "StatusDetail";           E = { $_.Status.FailureReason }},
            @{N = "RiskLevel";              E = { $_.RiskLevelDuringSignIn }},
            @{N = "RiskState";              E = { $_.RiskState }},
            @{N = "MfaDetail";              E = { $_.MfaDetail.AuthMethod }},
            @{N = "ConditionalAccessStatus"; E = { $_.ConditionalAccessStatus }},
            @{N = "CreatedDateTime";        E = { $_.CreatedDateTime }},
            @{N = "IsInteractive";          E = { $_.IsInteractive }},
            @{N = "CorrelationId";          E = { $_.CorrelationId }}
    }
}

function Get-LegacyAuthData {
    param($SignInLogs)
    if (-not $SignInLogs) { return $null }

    return Invoke-SafeCollection -CollectionName "Legacy Authentication Analysis" -ScriptBlock {
        $legacyClients = @("imap", "pop3", "smtp", "exchange activesync", "other clients",
                           "authenticated smtp", "exchange web services", "autodiscover")

        @($SignInLogs) | Where-Object {
            $client = $_.ClientAppUsed
            if ([string]::IsNullOrEmpty($client)) { return $false }
            $clientLower = $client.ToLower()
            $legacyClients | Where-Object { $clientLower -like "*$_*" }
        } | Select-Object UPN, DisplayName, AppDisplayName, ClientAppUsed,
            IPAddress, Location, Status, CreatedDateTime, RiskLevel
    }
}

function Get-TransportConfigData {
    return Invoke-SafeCollection -CollectionName "Transport Configuration EXO" -ScriptBlock {
        $transport = Get-TransportConfig -ErrorAction Stop
        [PSCustomObject]@{
            SmtpClientAuthenticationDisabled = $transport.SmtpClientAuthenticationDisabled
            TLSReceiveDomainSecureList       = ($transport.TLSReceiveDomainSecureList -join "; ")
            TLSSendDomainSecureList          = ($transport.TLSSendDomainSecureList -join "; ")
            MaxReceiveSize                   = $transport.MaxReceiveSize
            MaxSendSize                      = $transport.MaxSendSize
        }
    }
}

function Get-CASMailboxData {
    return Invoke-SafeCollection -CollectionName "EXO CAS Mailbox Settings" -ScriptBlock {
        # Get-EXOCASMailbox e' la versione REST moderna di Get-CASMailbox
        $casMailboxes = Get-EXOCASMailbox -ResultSize Unlimited -ErrorAction Stop
        $casMailboxes | Select-Object `
            @{N = "UPN";                             E = { $_.PrimarySmtpAddress }},
            @{N = "DisplayName";                     E = { $_.DisplayName }},
            @{N = "SmtpClientAuthenticationDisabled"; E = { $_.SmtpClientAuthenticationDisabled }},
            @{N = "PopEnabled";                      E = { $_.PopEnabled }},
            @{N = "ImapEnabled";                     E = { $_.ImapEnabled }},
            @{N = "ActiveSyncEnabled";               E = { $_.ActiveSyncEnabled }},
            @{N = "OWAEnabled";                      E = { $_.OWAEnabled }},
            @{N = "EwsEnabled";                      E = { $_.EwsEnabled }},
            @{N = "MapiEnabled";                     E = { $_.MapiEnabled }}
    }
}

function Get-MailboxData {
    return Invoke-SafeCollection -CollectionName "Mailboxes" -ScriptBlock {
        $mailboxes = Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop
        $mailboxes | Select-Object `
            @{N = "UPN";                        E = { $_.UserPrincipalName }},
            @{N = "DisplayName";                E = { $_.DisplayName }},
            @{N = "PrimarySmtp";                E = { $_.PrimarySmtpAddress }},
            @{N = "RecipientTypeDetails";       E = { $_.RecipientTypeDetails }},
            @{N = "ForwardingAddress";          E = { $_.ForwardingAddress }},
            @{N = "ForwardingSmtpAddress";      E = { $_.ForwardingSmtpAddress }},
            @{N = "DeliverToMailboxAndForward"; E = { $_.DeliverToMailboxAndForward }},
            @{N = "HiddenFromAddressLists";     E = { $_.HiddenFromAddressListsEnabled }},
            @{N = "WhenCreated";                E = { $_.WhenCreated }}
    }
}

function Get-DirectoryRolesData {
    return Invoke-SafeCollection -CollectionName "Directory Roles & Members" -ScriptBlock {
        $roles    = Get-MgDirectoryRole -All -ErrorAction Stop
        $roleData = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach ($role in $roles) {
            try {
                $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction Stop
                foreach ($member in $members) {
                    $userDetails  = $null
                    $memberOdType = $member.AdditionalProperties.'@odata.type'
                    try {
                        if ($memberOdType -eq '#microsoft.graph.user') {
                            $userDetails = Get-MgUser -UserId $member.Id `
                                -Property "UserPrincipalName,DisplayName,AccountEnabled,UserType" `
                                -ErrorAction SilentlyContinue
                        }
                    } catch { }

                    $memberUPN  = if ($userDetails -and $userDetails.UserPrincipalName) { $userDetails.UserPrincipalName } else { "N/A" }
                    $memberName = if ($userDetails -and $userDetails.DisplayName) {
                        $userDetails.DisplayName
                    } elseif ($member.AdditionalProperties.displayName) {
                        $member.AdditionalProperties.displayName
                    } else { "N/A" }
                    $acctEnabled = if ($null -ne $userDetails) { $userDetails.AccountEnabled } else { "N/A" }
                    $userType    = if ($userDetails -and $userDetails.UserType) { $userDetails.UserType } else { "N/A" }

                    $roleData.Add([PSCustomObject]@{
                        RoleName          = $role.DisplayName
                        RoleId            = $role.Id
                        MemberId          = $member.Id
                        MemberType        = $memberOdType
                        MemberUPN         = $memberUPN
                        MemberDisplayName = $memberName
                        AccountEnabled    = $acctEnabled
                        UserType          = $userType
                    })
                }
            } catch {
                Write-Log "Impossibile ottenere membri per ruolo '$($role.DisplayName)': $($_.Exception.Message)" "WARN"
            }
        }
        return $roleData.ToArray()
    }
}

function Get-OAuthGrantsData {
    # Cattura variabili di scope esterno per compatibilita con ThreadJob
    $highRiskOAuthScopes = $Script:HighRiskOAuthScopes

    return Invoke-SafeCollection -CollectionName "OAuth Grants & Service Principals" -ScriptBlock {
        # Nota: quando eseguito inline, $using: non e necessario; il blocco usa
        # il valore catturato in $highRiskOAuthScopes dalla chiusura padre.
        $localHighRiskScopes = if ($null -ne $using:highRiskOAuthScopes) {
            $using:highRiskOAuthScopes
        } else {
            @("Mail.ReadWrite","Mail.Send","Files.ReadWrite.All","Directory.ReadWrite.All")
        }

        $grants  = Get-MgOauth2PermissionGrant -All -ErrorAction Stop
        $spCache = @{}

        $grantData = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($grant in $grants) {
            # Gestione difensiva: salta grant con struttura anomala
            try {
                if ($null -eq $grant -or [string]::IsNullOrEmpty($grant.ClientId)) {
                    continue
                }

                # Lookup Service Principal con cache
                if (-not $spCache.ContainsKey($grant.ClientId)) {
                    try {
                        $sp = Get-MgServicePrincipal -ServicePrincipalId $grant.ClientId -ErrorAction SilentlyContinue
                        $spCache[$grant.ClientId] = $sp
                    } catch {
                        $spCache[$grant.ClientId] = $null
                    }
                }
                $sp = $spCache[$grant.ClientId]

                # DisplayName con fallback
                $spName = "N/A"
                try {
                    if ($sp -and $sp.PSObject.Properties.Name -contains 'DisplayName' -and $sp.DisplayName) {
                        $spName = $sp.DisplayName
                    }
                } catch { }

                # PublisherName con ricerca difensiva in piu proprieta
                $spPublisher = "N/A"
                try {
                    if ($sp) {
                        if ($sp.PSObject.Properties.Name -contains 'PublisherName' -and
                            $null -ne $sp.PublisherName -and $sp.PublisherName -ne '') {
                            $spPublisher = $sp.PublisherName
                        } elseif ($sp.AdditionalProperties -and
                                  $sp.AdditionalProperties -is [System.Collections.IDictionary] -and
                                  $sp.AdditionalProperties.ContainsKey('publisherName') -and
                                  $sp.AdditionalProperties['publisherName']) {
                            $spPublisher = $sp.AdditionalProperties['publisherName']
                        }
                    }
                } catch { }

                # Scopes e rischio
                $scopeString = ""
                try {
                    if ($grant.PSObject.Properties.Name -contains 'Scope' -and $grant.Scope) {
                        $scopeString = $grant.Scope
                    }
                } catch { }

                $scopes         = if ($scopeString) { $scopeString -split " " } else { @() }
                $highRiskScopes = @($scopes | Where-Object { $_ -in $localHighRiskScopes })
                $riskLevel      = if ($highRiskScopes.Count -gt 0) { "ALTO" } else { "BASSO" }

                # ExpiryTime con accesso difensivo
                $expiryTime = $null
                try {
                    if ($grant.PSObject.Properties.Name -contains 'ExpiryTime') {
                        $expiryTime = $grant.ExpiryTime
                    } elseif ($grant.AdditionalProperties -and
                              $grant.AdditionalProperties -is [System.Collections.IDictionary] -and
                              $grant.AdditionalProperties.ContainsKey('expiryTime')) {
                        $expiryTime = $grant.AdditionalProperties['expiryTime']
                    }
                } catch { }

                # ConsentType e PrincipalId con fallback
                $consentType = try { $grant.ConsentType } catch { "N/A" }
                $principalId = try { $grant.PrincipalId } catch { $null }
                $resourceId  = try { $grant.ResourceId  } catch { $null }

                $grantData.Add([PSCustomObject]@{
                    ClientId          = $grant.ClientId
                    ClientDisplayName = $spName
                    ClientPublisher   = $spPublisher
                    ConsentType       = $consentType
                    PrincipalId       = $principalId
                    ResourceId        = $resourceId
                    Scopes            = $scopeString
                    HighRiskScopes    = ($highRiskScopes -join "; ")
                    RiskLevel         = $riskLevel
                    ExpiryTime        = $expiryTime
                })
            } catch {
                # Grant con struttura anomala: skippa e continua
                $grantId = try { $grant.ClientId } catch { "sconosciuto" }
                Write-Warning "OAuth Grant '$grantId' saltato per struttura anomala: $($_.Exception.Message)"
            }
        }
        return $grantData.ToArray()
    }
}

function Get-RiskyUsersData {
    return Invoke-SafeCollection -CollectionName "Risky Users (richiede licenza P2)" -ScriptBlock {
        try {
            $riskyUsers = Get-MgRiskyUser -All -ErrorAction Stop
            $riskyUsers | Select-Object `
                @{N = "UPN";            E = { $_.UserPrincipalName }},
                @{N = "DisplayName";    E = { $_.UserDisplayName }},
                @{N = "RiskLevel";      E = { $_.RiskLevel }},
                @{N = "RiskState";      E = { $_.RiskState }},
                @{N = "RiskDetail";     E = { $_.RiskDetail }},
                @{N = "RiskLastUpdated"; E = { $_.RiskLastUpdatedDateTime }},
                @{N = "IsDeleted";      E = { $_.IsDeleted }},
                @{N = "IsProcessing";   E = { $_.IsProcessing }}
        } catch {
            $msg = $_.Exception.Message
            if ($msg -like "*Unauthorized*" -or $msg -like "*license*" -or
                $msg -like "*403*" -or $msg -like "*AAD Premium*") {
                Write-Log "Risky Users non disponibile (licenza P2 richiesta) - continuando..." "WARN"
                return @()
            }
            throw
        }
    }
}

function Get-RiskDetectionsData {
    # NOTA: Get-MgRiskDetection e' il cmdlet corretto nel modulo Microsoft.Graph >= 2.x
    # (Get-MgIdentityRiskDetection e' stato rimosso/rinominato).
    # Richiede licenza Azure AD P1 (base) o P2 (completa); gestisce gracefully l'assenza.
    return Invoke-SafeCollection -CollectionName "Risk Detections (richiede licenza P1/P2)" -ScriptBlock {
        try {
            $startDate  = (Get-Date).AddDays(-$LookbackDays).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
            $detections = Get-MgRiskDetection -Filter "detectedDateTime ge $startDate" -All -ErrorAction Stop

            $detections | Select-Object `
                @{N = "UPN";              E = { $_.UserPrincipalName }},
                @{N = "DisplayName";      E = { $_.UserDisplayName }},
                @{N = "DetectionType";    E = { $_.DetectionTimingType }},
                @{N = "RiskType";         E = { $_.RiskType }},
                @{N = "RiskLevel";        E = { $_.RiskLevel }},
                @{N = "RiskState";        E = { $_.RiskState }},
                @{N = "RiskDetail";       E = { $_.RiskDetail }},
                @{N = "Source";           E = { $_.Source }},
                @{N = "IPAddress";        E = { $_.IpAddress }},
                @{N = "Location";         E = { if ($_.Location) { "$($_.Location.City), $($_.Location.CountryOrRegion)" } else { "N/A" } }},
                @{N = "DetectedDateTime"; E = { $_.DetectedDateTime }},
                @{N = "LastUpdated";      E = { $_.LastUpdatedDateTime }}
        } catch {
            $msg = $_.Exception.Message
            # Gestisce sia l'assenza di licenza P1/P2 che errori di autorizzazione
            if ($msg -like "*Unauthorized*"    -or
                $msg -like "*license*"         -or
                $msg -like "*403*"             -or
                $msg -like "*AAD Premium*"     -or
                $msg -like "*AadPremiumLicense*" -or
                $msg -like "*Identity Protection*" -or
                $msg -like "*does not have a license*" -or
                $msg -like "*InsufficientComplexLicense*") {
                Write-Log "Risk Detections non disponibile (licenza Azure AD P1/P2 richiesta) - continuando..." "WARN"
                return @()
            }
            throw
        }
    }
}

function Get-MailboxRulesData {
    param($Mailboxes)
    if (-not $Mailboxes) { return $null }

    $mbxArray  = @($Mailboxes)
    $total     = $mbxArray.Count
    Write-Log "Raccolta: Inbox Rules & Forwarding per $total mailbox..." "INFO"

    $rulesData = [System.Collections.Generic.List[PSCustomObject]]::new()
    $fwdData   = [System.Collections.Generic.List[PSCustomObject]]::new()
    $processed = 0

    foreach ($mbx in $mbxArray) {
        $processed++
        $pct = [math]::Round(($processed / $total) * 100)
        Write-Progress -Activity "Analisi Inbox Rules & Forwarding" `
            -Status "Mailbox $processed/$total : $($mbx.UPN)" `
            -PercentComplete $pct

        if ($processed % 50 -eq 0) {
            Write-Log "  Progress Rules: $processed/$total mailbox analizzate..." "INFO"
        }

        # Forwarding dalle proprieta mailbox
        if ($mbx.ForwardingSmtpAddress -or $mbx.ForwardingAddress) {
            $isExternal = $mbx.ForwardingSmtpAddress -and
                          $mbx.ForwardingSmtpAddress -notlike "*$TenantName*"
            $fwdData.Add([PSCustomObject]@{
                UPN                        = $mbx.UPN
                DisplayName                = $mbx.DisplayName
                ForwardingAddress          = $mbx.ForwardingAddress
                ForwardingSmtpAddress      = $mbx.ForwardingSmtpAddress
                DeliverToMailboxAndForward = $mbx.DeliverToMailboxAndForward
                Source                     = "MailboxProperty"
                RiskLevel                  = if ($isExternal) { "ALTO" } else { "MEDIO" }
            })
        }

        # Inbox Rules
        try {
            $rules = Get-InboxRule -Mailbox $mbx.UPN -ErrorAction Stop
            foreach ($rule in $rules) {
                $isRisky = $rule.ForwardTo -or $rule.ForwardAsAttachmentTo -or
                           $rule.RedirectTo -or $rule.DeleteMessage -or
                           $rule.SoftDeleteMessage

                $rulesData.Add([PSCustomObject]@{
                    UPN                  = $mbx.UPN
                    DisplayName          = $mbx.DisplayName
                    RuleName             = $rule.Name
                    RuleEnabled          = $rule.Enabled
                    Priority             = $rule.Priority
                    ForwardTo            = ($rule.ForwardTo -join "; ")
                    ForwardAsAttachment  = ($rule.ForwardAsAttachmentTo -join "; ")
                    RedirectTo           = ($rule.RedirectTo -join "; ")
                    DeleteMessage        = $rule.DeleteMessage
                    SoftDeleteMessage    = $rule.SoftDeleteMessage
                    MoveToFolder         = $rule.MoveToFolder
                    SubjectContainsWords = ($rule.SubjectContainsWords -join "; ")
                    FromAddressContains  = ($rule.FromAddressContainsWords -join "; ")
                    RiskLevel            = if ($isRisky) { "ALTO" } else { "BASSO" }
                    Description          = $rule.Description
                })
            }
        } catch {
            Write-Log "  Impossibile leggere regole per $($mbx.UPN): $($_.Exception.Message)" "WARN"
        }
    }

    Write-Progress -Activity "Analisi Inbox Rules & Forwarding" -Completed
    Write-Log "Completato: Inbox Rules - $($rulesData.Count) regole, $($fwdData.Count) forward trovati." "SUCCESS"
    return @{
        Rules      = $rulesData.ToArray()
        Forwarding = $fwdData.ToArray()
    }
}

#endregion

#region ─── CALCOLO SECURITY SCORE ────────────────────────────────────────────

function Get-SecurityScore {
    param(
        $TransportConfig,
        $MfaData,
        $CaPolicies,
        $DirectoryRoles,
        $CasMailboxes,
        $OAuthGrants
    )

    $score        = 0
    $scoreDetails = [System.Collections.Generic.List[PSCustomObject]]::new()
    $maxScore     = ($Script:ScoreThresholds.Values | Measure-Object -Sum).Sum

    # 1. SMTP AUTH Globale
    $smtpDisabled = ($TransportConfig -and $TransportConfig.SmtpClientAuthenticationDisabled -eq $true)
    $pts = if ($smtpDisabled) { $Script:ScoreThresholds.SmtpAuthDisabled } else { 0 }
    $score += $pts
    $scoreDetails.Add([PSCustomObject]@{
        Controllo = "SMTP AUTH Globale Disabilitato"
        Punteggio = $pts
        Massimo   = $Script:ScoreThresholds.SmtpAuthDisabled
        Stato     = if ($smtpDisabled) { "OK" } else { "RISCHIO" }
        Livello   = if ($smtpDisabled) { "OK" } else { "ALTO" }
        Dettaglio = if ($smtpDisabled) { "SMTP AUTH disabilitato a livello tenant" } else { "SMTP AUTH abilitato - rischio autenticazione legacy" }
    })

    # 2. MFA Coverage
    $mfaTotal      = 0
    $mfaRegistered = 0
    $mfaCoverage   = 0
    if ($MfaData -and @($MfaData).Count -gt 0) {
        $mfaTotal      = @($MfaData).Count
        $mfaRegistered = @($MfaData | Where-Object { $_.IsMfaRegistered -eq $true }).Count
        $mfaCoverage   = if ($mfaTotal -gt 0) { [math]::Round(($mfaRegistered / $mfaTotal) * 100, 1) } else { 0 }
    }
    $mfaStats = @{ Registered = $mfaRegistered; Total = $mfaTotal; Coverage = $mfaCoverage }
    $mfaPts   = if ($mfaCoverage -ge 90) { $Script:ScoreThresholds.MfaHighCoverage } `
                elseif ($mfaCoverage -ge 70) { $Script:ScoreThresholds.MfaMedCoverage } else { 0 }
    $score += $mfaPts
    $mfaLevel = if ($mfaCoverage -ge 90) { "OK" } elseif ($mfaCoverage -ge 70) { "MEDIO" } else { "CRITICO" }
    $scoreDetails.Add([PSCustomObject]@{
        Controllo = "Copertura MFA Utenti"
        Punteggio = $mfaPts
        Massimo   = $Script:ScoreThresholds.MfaHighCoverage
        Stato     = if ($mfaPts -eq $Script:ScoreThresholds.MfaHighCoverage) { "OK" } else { "RISCHIO" }
        Livello   = $mfaLevel
        Dettaglio = "Coverage: $mfaCoverage% ($mfaRegistered/$mfaTotal utenti)"
    })

    # 3. CA con MFA
    $caMfaPolicies = @()
    if ($CaPolicies) {
        $caMfaPolicies = @($CaPolicies | Where-Object {
            $_.State -eq "enabled" -and $_.GrantControls -like "*mfa*"
        })
    }
    $caMfaPts = if ($caMfaPolicies.Count -gt 0) { $Script:ScoreThresholds.CaMfaEnabled } else { 0 }
    $score += $caMfaPts
    $scoreDetails.Add([PSCustomObject]@{
        Controllo = "Conditional Access con MFA"
        Punteggio = $caMfaPts
        Massimo   = $Script:ScoreThresholds.CaMfaEnabled
        Stato     = if ($caMfaPolicies.Count -gt 0) { "OK" } else { "RISCHIO" }
        Livello   = if ($caMfaPolicies.Count -gt 0) { "OK" } else { "CRITICO" }
        Dettaglio = if ($caMfaPolicies.Count -gt 0) { "$($caMfaPolicies.Count) policy CA con MFA attiva" } else { "Nessuna policy CA con MFA trovata" }
    })

    # 4. Block Legacy Auth via CA
    $blockLegacyPolicies = @()
    if ($CaPolicies) {
        $blockLegacyPolicies = @($CaPolicies | Where-Object {
            $_.State -eq "enabled" -and
            ($_.ClientAppTypes -like "*exchangeActiveSync*" -or $_.ClientAppTypes -like "*other*") -and
            $_.GrantControls -like "*block*"
        })
    }
    $blockLegacyPts = if ($blockLegacyPolicies.Count -gt 0) { $Script:ScoreThresholds.BlockLegacyAuth } else { 0 }
    $score += $blockLegacyPts
    $scoreDetails.Add([PSCustomObject]@{
        Controllo = "Blocco Legacy Auth via CA"
        Punteggio = $blockLegacyPts
        Massimo   = $Script:ScoreThresholds.BlockLegacyAuth
        Stato     = if ($blockLegacyPolicies.Count -gt 0) { "OK" } else { "RISCHIO" }
        Livello   = if ($blockLegacyPolicies.Count -gt 0) { "OK" } else { "ALTO" }
        Dettaglio = if ($blockLegacyPolicies.Count -gt 0) { "Policy CA attiva per blocco legacy auth" } else { "Nessuna policy CA blocca autenticazione legacy" }
    })

    # 5. Global Admin Count
    $globalAdmins = @()
    if ($DirectoryRoles) {
        $globalAdmins = @($DirectoryRoles | Where-Object { $_.RoleName -eq "Global Administrator" })
    }
    $adminCount = $globalAdmins.Count
    $adminPts   = if ($adminCount -le 5 -and $adminCount -ge 2) { $Script:ScoreThresholds.FewGlobalAdmins } else { 0 }
    $adminLevel = if ($adminCount -eq 0) { "CRITICO" } elseif ($adminCount -gt 10) { "ALTO" } elseif ($adminCount -gt 5) { "MEDIO" } else { "OK" }
    $score += $adminPts
    $scoreDetails.Add([PSCustomObject]@{
        Controllo = "Numero Global Administrators"
        Punteggio = $adminPts
        Massimo   = $Script:ScoreThresholds.FewGlobalAdmins
        Stato     = if ($adminPts -gt 0) { "OK" } else { "RISCHIO" }
        Livello   = $adminLevel
        Dettaglio = "$adminCount Global Admin trovati (ottimale: 2-5)"
    })

    # 6. POP/IMAP Exposure
    $popImapExposed = 0
    $casCount       = 0
    if ($CasMailboxes) {
        $casCount       = @($CasMailboxes).Count
        $popImapExposed = @($CasMailboxes | Where-Object { $_.PopEnabled -eq $true -or $_.ImapEnabled -eq $true }).Count
    }
    $popImapPct = if ($casCount -gt 0) { [math]::Round(($popImapExposed / $casCount) * 100, 1) } else { 0 }
    $popImapPts = if ($popImapPct -lt 5) { $Script:ScoreThresholds.PoPImapDisabled } else { 0 }
    $score += $popImapPts
    $scoreDetails.Add([PSCustomObject]@{
        Controllo = "Esposizione POP3/IMAP"
        Punteggio = $popImapPts
        Massimo   = $Script:ScoreThresholds.PoPImapDisabled
        Stato     = if ($popImapPts -gt 0) { "OK" } else { "RISCHIO" }
        Livello   = if ($popImapPct -lt 5) { "OK" } elseif ($popImapPct -lt 20) { "MEDIO" } else { "ALTO" }
        Dettaglio = "$popImapExposed mailbox con POP/IMAP abilitato ($popImapPct%)"
    })

    # 7. OAuth High-Risk Scopes
    $highRiskOAuth = 0
    if ($OAuthGrants) {
        $highRiskOAuth = @($OAuthGrants | Where-Object { $_.RiskLevel -eq "ALTO" }).Count
    }
    $oauthPts = if ($highRiskOAuth -eq 0) { $Script:ScoreThresholds.LowOauthRisk } else { 0 }
    $score += $oauthPts
    $scoreDetails.Add([PSCustomObject]@{
        Controllo = "Scope OAuth ad Alto Rischio"
        Punteggio = $oauthPts
        Massimo   = $Script:ScoreThresholds.LowOauthRisk
        Stato     = if ($oauthPts -gt 0) { "OK" } else { "RISCHIO" }
        Livello   = if ($highRiskOAuth -eq 0) { "OK" } elseif ($highRiskOAuth -le 5) { "MEDIO" } else { "ALTO" }
        Dettaglio = "$highRiskOAuth grant OAuth con scope ad alto rischio"
    })

    $percentage = if ($maxScore -gt 0) { [math]::Round(($score / $maxScore) * 100, 1) } else { 0 }
    $rating     = if ($percentage -ge 80) { "BUONO" } `
                  elseif ($percentage -ge 60) { "SUFFICIENTE" } `
                  elseif ($percentage -ge 40) { "INSUFFICIENTE" } else { "CRITICO" }

    return @{
        Score        = $score
        MaxScore     = $maxScore
        Percentage   = $percentage
        Rating       = $rating
        Details      = $scoreDetails.ToArray()
        MfaStats     = $mfaStats
        AdminCount   = $adminCount
        GlobalAdmins = $globalAdmins
    }
}

#endregion

#region ─── EXPORT CSV ─────────────────────────────────────────────────────────

function Export-DataToCSV {
    param(
        [string]$CsvDir,
        $MfaData,
        $CaData,
        $SignInData,
        $LegacyData,
        $TransportData,
        $CasData,
        $MailboxData,
        $RolesData,
        $OAuthData,
        $RiskyUsersData,
        $RiskDetectionsData,
        $RulesData
    )

    Write-Section "EXPORT CSV"
    $exported = [System.Collections.Generic.List[string]]::new()

    $rulesArr   = if ($RulesData -and $RulesData.ContainsKey("Rules"))      { $RulesData["Rules"] }      else { $null }
    $fwdArr     = if ($RulesData -and $RulesData.ContainsKey("Forwarding")) { $RulesData["Forwarding"] } else { $null }
    $transportArr = if ($TransportData) { @($TransportData) } else { $null }

    $csvMap = @(
        @{ Name = "MFA_Registration";  Data = $MfaData }
        @{ Name = "Conditional_Access"; Data = $CaData }
        @{ Name = "SignIn_Logs";        Data = $SignInData }
        @{ Name = "Legacy_Auth";        Data = $LegacyData }
        @{ Name = "Transport_Config";   Data = $transportArr }
        @{ Name = "CAS_Mailboxes";      Data = $CasData }
        @{ Name = "Mailboxes";          Data = $MailboxData }
        @{ Name = "Directory_Roles";    Data = $RolesData }
        @{ Name = "OAuth_Grants";       Data = $OAuthData }
        @{ Name = "Risky_Users";        Data = $RiskyUsersData }
        @{ Name = "Risk_Detections";    Data = $RiskDetectionsData }
        @{ Name = "Inbox_Rules";        Data = $rulesArr }
        @{ Name = "Forwarding";         Data = $fwdArr }
    )

    # CSV sempre generati anche se vuoti (garantisce presenza file per post-processing)
    $alwaysExport = @("OAuth_Grants")

    $csvIndex = 0
    foreach ($item in $csvMap) {
        $csvIndex++
        $pct     = [math]::Round(($csvIndex / $csvMap.Count) * 100)
        $csvPath = Join-Path $CsvDir "$($item.Name).csv"
        Write-Progress -Activity "Export CSV" -Status "$($item.Name).csv" -PercentComplete $pct

        $hasData = ($item.Data -and @($item.Data).Count -gt 0)
        $mustExport = ($item.Name -in $alwaysExport)

        if ($hasData) {
            try {
                @($item.Data) | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
                $count = @($item.Data).Count
                Write-Log "  CSV: $($item.Name).csv ($count record)" "SUCCESS"
                $exported.Add($csvPath)
            } catch {
                Write-Log "  Errore export CSV $($item.Name): $($_.Exception.Message)" "ERROR"
            }
        } elseif ($mustExport) {
            # Genera file vuoto con header placeholder per garantire la presenza del file
            try {
                [PSCustomObject]@{
                    Nota = "Nessun dato raccolto - grant OAuth assenti o raccolta saltata per timeout/errore"
                } | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
                Write-Log "  CSV: $($item.Name).csv (vuoto - file garantito)" "WARN"
                $exported.Add($csvPath)
            } catch {
                Write-Log "  Errore export CSV vuoto $($item.Name): $($_.Exception.Message)" "ERROR"
            }
        } else {
            Write-Log "  Skip CSV: $($item.Name) (nessun dato)" "WARN"
        }
    }
    Write-Progress -Activity "Export CSV" -Completed
    return $exported.ToArray()
}

#endregion

#region ─── HELPER EXCEL (ImportExcel) ────────────────────────────────────────

function Add-ExcelTitleRow {
    param(
        [OfficeOpenXml.ExcelWorksheet]$Worksheet,
        [int]$Row,
        [string]$Title,
        [int]$MergeEnd = 8
    )
    $Worksheet.Cells[$Row, 1, $Row, $MergeEnd].Merge = $true
    $Worksheet.Cells[$Row, 1].Value = $Title
    $Worksheet.Cells[$Row, 1].Style.Font.Bold = $true
    $Worksheet.Cells[$Row, 1].Style.Font.Size = 14
    $Worksheet.Cells[$Row, 1].Style.Font.Color.SetColor(
        [System.Drawing.ColorTranslator]::FromHtml("#$($Script:Colors.TitleFg)"))
    $Worksheet.Cells[$Row, 1].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $Worksheet.Cells[$Row, 1].Style.Fill.BackgroundColor.SetColor(
        [System.Drawing.ColorTranslator]::FromHtml("#$($Script:Colors.TitleBg)"))
    $Worksheet.Cells[$Row, 1].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $Worksheet.Row($Row).Height = 30
}

function Add-ExcelHeaderRow {
    param(
        [OfficeOpenXml.ExcelWorksheet]$Worksheet,
        [int]$Row,
        [string[]]$Headers,
        [string]$BgColor = $Script:Colors.HeaderBg,
        [string]$FgColor = $Script:Colors.HeaderFg
    )
    for ($i = 0; $i -lt $Headers.Count; $i++) {
        $cell = $Worksheet.Cells[$Row, ($i + 1)]
        $cell.Value = $Headers[$i]
        $cell.Style.Font.Bold = $true
        $cell.Style.Font.Color.SetColor(
            [System.Drawing.ColorTranslator]::FromHtml("#$FgColor"))
        $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $cell.Style.Fill.BackgroundColor.SetColor(
            [System.Drawing.ColorTranslator]::FromHtml("#$BgColor"))
        $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
        $cell.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
        $cell.Style.WrapText            = $true
        $cell.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
    }
}

function Set-ExcelCellColor {
    param(
        [OfficeOpenXml.ExcelWorksheet]$Worksheet,
        [int]$Row,
        [int]$Col,
        [string]$BgColor,
        [string]$FgColor = "000000"
    )
    $cell = $Worksheet.Cells[$Row, $Col]
    $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $cell.Style.Fill.BackgroundColor.SetColor(
        [System.Drawing.ColorTranslator]::FromHtml("#$BgColor"))
    $cell.Style.Font.Color.SetColor(
        [System.Drawing.ColorTranslator]::FromHtml("#$FgColor"))
}

function Set-ExcelRowColor {
    param(
        [OfficeOpenXml.ExcelWorksheet]$Worksheet,
        [int]$Row,
        [int]$ColCount,
        [string]$BgColor,
        [string]$FgColor = "000000"
    )
    for ($c = 1; $c -le $ColCount; $c++) {
        Set-ExcelCellColor -Worksheet $Worksheet -Row $Row -Col $c -BgColor $BgColor -FgColor $FgColor
    }
}

function Set-ExcelColumnWidths {
    param(
        [OfficeOpenXml.ExcelWorksheet]$Worksheet,
        [int[]]$Widths
    )
    for ($i = 0; $i -lt $Widths.Count; $i++) {
        $Worksheet.Column($i + 1).Width = $Widths[$i]
    }
}

function Get-RiskColors {
    param([string]$Level)
    switch ($Level) {
        "CRITICO" { return @($Script:Colors.CriticalBg, $Script:Colors.CriticalFg) }
        "ALTO"    { return @($Script:Colors.HighBg, $Script:Colors.HighFg) }
        "MEDIO"   { return @($Script:Colors.MediumBg, $Script:Colors.MediumFg) }
        "OK"      { return @($Script:Colors.OkBg, $Script:Colors.OkFg) }
        default   { return @($Script:Colors.InfoBg, $Script:Colors.InfoFg) }
    }
}

#endregion

#region ─── GENERAZIONE EXCEL ──────────────────────────────────────────────────

function New-ExcelWorkbook {
    param(
        [string]$ExcelPath,
        $ScoreResult,
        $MfaData,
        $CaData,
        $SignInData,
        $LegacyData,
        $TransportData,
        $CasData,
        $MailboxData,
        $RolesData,
        $OAuthData,
        $RiskyUsersData,
        $RiskDetectionsData,
        $RulesData
    )

    Write-Section "GENERAZIONE WORKBOOK EXCEL"

    if (Test-Path $ExcelPath) { Remove-Item $ExcelPath -Force }

    # Usa Open-ExcelPackage di ImportExcel per creare/aprire il package.
    # ImportExcel carica l'assembly EPPlus (OfficeOpenXml) come dipendenza,
    # rendendo disponibili i tipi [OfficeOpenXml.*] usati dalle funzioni helper.
    # Questo e' il modo raccomandato per garantire la compatibilita di versione.
    $excelPkg = Open-ExcelPackage -Path $ExcelPath -Create

    # ─── Helper locale: aggiunge un foglio e disabilita le griglie ───────────
    function New-Sheet {
        param([string]$Name)
        $ws = $excelPkg.Workbook.Worksheets.Add($Name)
        $ws.View.ShowGridLines = $false
        return $ws
    }

    #─── 1. SINTESI ESECUTIVA ─────────────────────────────────────────────────
    Write-Log "  Creazione: Sintesi Esecutiva..." "INFO"
    $ws1 = New-Sheet "Sintesi Esecutiva"

    Add-ExcelTitleRow -Worksheet $ws1 -Row 1 `
        -Title "VALUTAZIONE SICUREZZA MICROSOFT 365 - SINTESI ESECUTIVA" -MergeEnd 8

    $row = 3

    $incidentInfo = if (-not [string]::IsNullOrEmpty($IncidentUser)) {
        $IncidentUser
    } else {
        "N/A (assessment generale)"
    }

    $infoItems = @(
        @("Tenant",          $TenantName),
        @("Data Assessment", (Get-Date -Format "dd/MM/yyyy HH:mm")),
        @("Periodo Analisi", "Ultimi $LookbackDays giorni"),
        @("Versione Script", $Script:AssessmentVersion),
        @("Utente Indagine", $incidentInfo)
    )
    foreach ($info in $infoItems) {
        $ws1.Cells[$row, 1].Value = $info[0]
        $ws1.Cells[$row, 1].Style.Font.Bold = $true
        $ws1.Cells[$row, 2, $row, 4].Merge = $true
        $ws1.Cells[$row, 2].Value = $info[1]
        $row++
    }
    $row++

    # Score box
    $ws1.Cells[$row, 1, $row, 8].Merge = $true
    $scoreLabel = "SECURITY SCORE: $($ScoreResult.Score)/$($ScoreResult.MaxScore) ($($ScoreResult.Percentage)%) - $($ScoreResult.Rating)"
    $ws1.Cells[$row, 1].Value = $scoreLabel
    $ws1.Cells[$row, 1].Style.Font.Bold = $true
    $ws1.Cells[$row, 1].Style.Font.Size = 18
    $ws1.Cells[$row, 1].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center

    switch ($scoreResult.Rating) {

    "BUONO" {
        $riskLevel = "OK"
    }

    "SUFFICIENTE" {
        $riskLevel = "MEDIO"
    }

    "INSUFFICIENTE" {
        $riskLevel = "ALTO"
    }

    default {
        $riskLevel = "CRITICO"
    }
}

$scoreColorPair = Get-RiskColors -Level $riskLevel
    Set-ExcelCellColor -Worksheet $ws1 -Row $row -Col 1 `
        -BgColor $scoreColorPair[0] -FgColor $scoreColorPair[1]
    $ws1.Row($row).Height = 40
    $row += 2

    # Riepilogo rischi
    Add-ExcelHeaderRow -Worksheet $ws1 -Row $row `
        -Headers @("Area", "Stato", "Livello Rischio", "Dettaglio", "Azione Raccomandata")
    $row++

    $caEnabledCount  = if ($CaData) { @($CaData | Where-Object { $_.State -eq "enabled" }).Count } else { 0 }
    $caTotalCount    = if ($CaData) { @($CaData).Count } else { 0 }
    $smtpDisabledVal = ($TransportData -and $TransportData.SmtpClientAuthenticationDisabled -eq $true)

    $riskSummary = @(
        @{
            Area   = "Autenticazione Multi-Fattore"
            Stato  = if ($ScoreResult.MfaStats.Coverage -ge 90) { "Buona copertura" } else { "Copertura insufficiente" }
            Livello= if ($ScoreResult.MfaStats.Coverage -ge 90) { "OK" } elseif ($ScoreResult.MfaStats.Coverage -ge 70) { "MEDIO" } else { "CRITICO" }
            Detail = "Coverage: $($ScoreResult.MfaStats.Coverage)%"
            Action = if ($ScoreResult.MfaStats.Coverage -ge 90) { "Monitorare" } else { "Abilitare MFA per tutti gli utenti" }
        },
        @{
            Area   = "SMTP Authentication"
            Stato  = if ($smtpDisabledVal) { "Disabilitato" } else { "Abilitato" }
            Livello= if ($smtpDisabledVal) { "OK" } else { "ALTO" }
            Detail = if ($smtpDisabledVal) { "SMTP AUTH disabilitato globalmente" } else { "SMTP AUTH abilitato - rischio credential spray" }
            Action = if ($smtpDisabledVal) { "Monitorare" } else { "Disabilitare SMTP AUTH, usare OAuth2" }
        },
        @{
            Area   = "Global Administrators"
            Stato  = if ($ScoreResult.AdminCount -le 5 -and $ScoreResult.AdminCount -ge 2) { "Numero adeguato" } else { "Numero non ottimale" }
            Livello= if ($ScoreResult.AdminCount -le 5) { "OK" } elseif ($ScoreResult.AdminCount -le 10) { "MEDIO" } else { "ALTO" }
            Detail = "$($ScoreResult.AdminCount) Global Admin (ottimale: 2-5)"
            Action = if ($ScoreResult.AdminCount -gt 5) { "Ridurre Global Admin, usare ruoli specifici" } else { "Verificare periodicamente" }
        },
        @{
            Area   = "Conditional Access"
            Stato  = "$caEnabledCount policy attive"
            Livello= if ($caEnabledCount -gt 0) { "OK" } else { "CRITICO" }
            Detail = "$caTotalCount policy totali"
            Action = "Implementare policy CA per MFA e blocco legacy auth"
        }
    )

    foreach ($risk in $riskSummary) {
        $ws1.Cells[$row, 1].Value = $risk.Area
        $ws1.Cells[$row, 2].Value = $risk.Stato
        $ws1.Cells[$row, 3].Value = $risk.Livello
        $ws1.Cells[$row, 4].Value = $risk.Detail
        $ws1.Cells[$row, 5].Value = $risk.Action
        $rcp = Get-RiskColors -Level $risk.Livello
        Set-ExcelRowColor -Worksheet $ws1 -Row $row -ColCount 5 -BgColor $rcp[0] -FgColor $rcp[1]
        $row++
    }
    Set-ExcelColumnWidths -Worksheet $ws1 -Widths @(30, 28, 16, 40, 45)

    #─── 2. SCORE SICUREZZA ──────────────────────────────────────────────────
    Write-Log "  Creazione: Score Sicurezza..." "INFO"
    $ws2 = New-Sheet "Score Sicurezza"

    Add-ExcelTitleRow -Worksheet $ws2 -Row 1 -Title "PUNTEGGIO DI SICUREZZA DETTAGLIATO" -MergeEnd 7

    $row = 3
    Add-ExcelHeaderRow -Worksheet $ws2 -Row $row `
        -Headers @("Controllo di Sicurezza", "Punteggio", "Massimo", "% Controllo", "Stato", "Livello Rischio", "Dettaglio")
    $row++

    foreach ($detail in $ScoreResult.Details) {
        $ws2.Cells[$row, 1].Value = $detail.Controllo
        $ws2.Cells[$row, 2].Value = $detail.Punteggio
        $ws2.Cells[$row, 3].Value = $detail.Massimo
        if ($detail.Massimo -gt 0) {
            $ws2.Cells[$row, 4].Value = "=B$row/C$row"
        } else {
            $ws2.Cells[$row, 4].Value = 0
        }
        $ws2.Cells[$row, 4].Style.Numberformat.Format = "0%"
        $ws2.Cells[$row, 5].Value = $detail.Stato
        $ws2.Cells[$row, 6].Value = $detail.Livello
        $ws2.Cells[$row, 7].Value = $detail.Dettaglio
        $rcp = Get-RiskColors -Level $detail.Livello
        Set-ExcelRowColor -Worksheet $ws2 -Row $row -ColCount 7 -BgColor $rcp[0] -FgColor $rcp[1]
        $row++
    }

    # Riga totale
    $row++
    $ws2.Cells[$row, 1].Value = "TOTALE"
    $ws2.Cells[$row, 1].Style.Font.Bold = $true
    $ws2.Cells[$row, 2].Value = $ScoreResult.Score
    $ws2.Cells[$row, 2].Style.Font.Bold = $true
    $ws2.Cells[$row, 3].Value = $ScoreResult.MaxScore
    $ws2.Cells[$row, 3].Style.Font.Bold = $true
    $ws2.Cells[$row, 4].Value = ($ScoreResult.Percentage / 100)
    $ws2.Cells[$row, 4].Style.Numberformat.Format = "0.0%"
    $ws2.Cells[$row, 4].Style.Font.Bold = $true
    $ws2.Cells[$row, 5].Value = "RATING: $($ScoreResult.Rating)"
    $ws2.Cells[$row, 5].Style.Font.Bold = $true
    Set-ExcelColumnWidths -Worksheet $ws2 -Widths @(42, 12, 12, 15, 18, 16, 60)

    #─── 3. AZIONI PRIORITARIE ───────────────────────────────────────────────
    Write-Log "  Creazione: Azioni Prioritarie..." "INFO"
    $ws3 = New-Sheet "Azioni Prioritarie"

    Add-ExcelTitleRow -Worksheet $ws3 -Row 1 -Title "PIANO DI AZIONE PRIORITARIO" -MergeEnd 7

    $row = 3
    Add-ExcelHeaderRow -Worksheet $ws3 -Row $row `
        -Headers @("Priorita", "Categoria", "Azione", "Impatto", "Sforzo", "Scadenza", "Note Tecniche")
    $row++

    $actionPlan = @(
        @{ P = "1 - CRITICA"; Cat = "MFA";         Action = "Abilitare MFA per tutti gli utenti senza registrazione";                      Impact = "CRITICO"; Effort = "BASSO";  Due = "Immediato"; Note = "Usare Conditional Access o MFA per utente" }
        @{ P = "1 - CRITICA"; Cat = "CA Policy";   Action = "Creare policy CA per richiedere MFA a tutti gli utenti";                       Impact = "CRITICO"; Effort = "BASSO";  Due = "7 giorni";  Note = "Escludere account break-glass" }
        @{ P = "1 - CRITICA"; Cat = "Legacy Auth"; Action = "Creare policy CA per bloccare autenticazione legacy (EAS, Other)";             Impact = "ALTO";    Effort = "BASSO";  Due = "14 giorni"; Note = "Testare con report-only prima di abilitare" }
        @{ P = "2 - ALTA";    Cat = "SMTP AUTH";   Action = "Disabilitare SMTP AUTH globalmente in Exchange Online";                        Impact = "ALTO";    Effort = "MEDIO";  Due = "30 giorni"; Note = "Verificare applicazioni che usano SMTP AUTH" }
        @{ P = "2 - ALTA";    Cat = "Admin";       Action = "Ridurre il numero di Global Administrator, usare ruoli specifici";             Impact = "ALTO";    Effort = "MEDIO";  Due = "30 giorni"; Note = "Usare Exchange Admin, SharePoint Admin, ecc." }
        @{ P = "2 - ALTA";    Cat = "OAuth";       Action = "Revisione e revoca grant OAuth con scope ad alto rischio";                     Impact = "ALTO";    Effort = "MEDIO";  Due = "30 giorni"; Note = "Focus su Mail.ReadWrite, Files.ReadWrite.All" }
        @{ P = "3 - MEDIA";   Cat = "POP/IMAP";    Action = "Disabilitare POP3 e IMAP per mailbox che non ne necessitano";                 Impact = "MEDIO";   Effort = "BASSO";  Due = "60 giorni"; Note = "Verificare client legacy che usano POP/IMAP" }
        @{ P = "3 - MEDIA";   Cat = "Forwarding";  Action = "Revisione e rimozione regole di forwarding esterno non autorizzate";          Impact = "ALTO";    Effort = "MEDIO";  Due = "30 giorni"; Note = "Focus su forward verso domini esterni" }
        @{ P = "3 - MEDIA";   Cat = "Monitoring";  Action = "Abilitare Microsoft Defender for Office 365 e Identity Protection";           Impact = "ALTO";    Effort = "ALTO";   Due = "60 giorni"; Note = "Richiede licenza P2 per funzionalita complete" }
        @{ P = "4 - BASSA";   Cat = "Admin";       Action = "Implementare Privileged Identity Management (PIM) per ruoli amministrativi";  Impact = "ALTO";    Effort = "ALTO";   Due = "90 giorni"; Note = "Richiede Azure AD P2" }
        @{ P = "4 - BASSA";   Cat = "Monitoring";  Action = "Configurare alert SIEM per eventi di sicurezza critici";                      Impact = "MEDIO";   Effort = "ALTO";   Due = "90 giorni"; Note = "Log Analytics, Microsoft Sentinel" }
        @{ P = "4 - BASSA";   Cat = "Access";      Action = "Abilitare Continuous Access Evaluation (CAE) per sessioni real-time";         Impact = "MEDIO";   Effort = "BASSO";  Due = "90 giorni"; Note = "Disponibile con licenza E3/E5" }
    )

    foreach ($action in $actionPlan) {
        $ws3.Cells[$row, 1].Value = $action.P
        $ws3.Cells[$row, 2].Value = $action.Cat
        $ws3.Cells[$row, 3].Value = $action.Action
        $ws3.Cells[$row, 4].Value = $action.Impact
        $ws3.Cells[$row, 5].Value = $action.Effort
        $ws3.Cells[$row, 6].Value = $action.Due
        $ws3.Cells[$row, 7].Value = $action.Note
        $level = switch -Wildcard ($action.P) {
            "1*" { "CRITICO" }
            "2*" { "ALTO"    }
            "3*" { "MEDIO"   }
            default { "OK"   }
        }
        $rcp = Get-RiskColors -Level $level
        Set-ExcelCellColor -Worksheet $ws3 -Row $row -Col 1 -BgColor $rcp[0] -FgColor $rcp[1]
        $row++
    }
    Set-ExcelColumnWidths -Worksheet $ws3 -Widths @(14, 14, 70, 12, 12, 14, 60)

    #─── 4. MFA COVERAGE ─────────────────────────────────────────────────────
    Write-Log "  Creazione: MFA Coverage..." "INFO"
    $ws4 = New-Sheet "MFA Coverage"

    Add-ExcelTitleRow -Worksheet $ws4 -Row 1 -Title "STATO REGISTRAZIONE MFA UTENTI" -MergeEnd 9

    if ($MfaData -and @($MfaData).Count -gt 0) {
        $mfaArr          = @($MfaData)
        $mfaTotale       = $mfaArr.Count
        $mfaConMfa       = @($mfaArr | Where-Object { $_.IsMfaRegistered -eq $true }).Count
        $mfaCapable      = @($mfaArr | Where-Object { $_.IsMfaCapable -eq $true }).Count
        $mfaPasswordless = @($mfaArr | Where-Object { $_.IsPasswordlessCapable -eq $true }).Count
        $adminConMfa     = @($mfaArr | Where-Object { $_.IsAdmin -eq $true -and $_.IsMfaRegistered -eq $true }).Count
        $adminSenzaMfa   = @($mfaArr | Where-Object { $_.IsAdmin -eq $true -and $_.IsMfaRegistered -ne $true }).Count

        $row = 3
        $statItems = @(
            @("Totale Utenti Analizzati",   $mfaTotale),
            @("Utenti con MFA Registrata",  $mfaConMfa),
            @("Utenti con MFA Capable",     $mfaCapable),
            @("Utenti Passwordless",        $mfaPasswordless),
            @("Admin con MFA",              $adminConMfa),
            @("Admin SENZA MFA",            $adminSenzaMfa),
            @("Coverage MFA %",             "$($ScoreResult.MfaStats.Coverage)%")
        )
        foreach ($stat in $statItems) {
            $ws4.Cells[$row, 1].Value = $stat[0]
            $ws4.Cells[$row, 1].Style.Font.Bold = $true
            $ws4.Cells[$row, 2].Value = $stat[1]
            if ($stat[0] -like "*SENZA*" -and ([int]$stat[1] -gt 0)) {
                Set-ExcelCellColor -Worksheet $ws4 -Row $row -Col 2 `
                    -BgColor $Script:Colors.CriticalBg -FgColor $Script:Colors.CriticalFg
            }
            $row++
        }
        $row++

        Add-ExcelHeaderRow -Worksheet $ws4 -Row $row `
            -Headers @("UPN", "Display Name", "Admin", "MFA Registrata", "MFA Capable", "Passwordless", "SSPR", "Metodo Default", "Metodi Registrati")
        $row++

        foreach ($user in ($mfaArr | Sort-Object IsAdmin -Descending)) {
            $ws4.Cells[$row, 1].Value = $user.UPN
            $ws4.Cells[$row, 2].Value = $user.DisplayName
            $ws4.Cells[$row, 3].Value = if ($user.IsAdmin)             { "Si" } else { "No" }
            $ws4.Cells[$row, 4].Value = if ($user.IsMfaRegistered)     { "Si" } else { "No" }
            $ws4.Cells[$row, 5].Value = if ($user.IsMfaCapable)        { "Si" } else { "No" }
            $ws4.Cells[$row, 6].Value = if ($user.IsPasswordlessCapable){ "Si" } else { "No" }
            $ws4.Cells[$row, 7].Value = if ($user.IsSsprRegistered)    { "Si" } else { "No" }
            $ws4.Cells[$row, 8].Value = $user.DefaultMfaMethod
            $ws4.Cells[$row, 9].Value = $user.MethodsRegistered

            if ($user.IsMfaRegistered -ne $true) {
                $bg = if ($user.IsAdmin -eq $true) { $Script:Colors.CriticalBg } else { $Script:Colors.HighBg }
                Set-ExcelRowColor -Worksheet $ws4 -Row $row -ColCount 9 -BgColor $bg -FgColor "FFFFFF"
            } elseif ($row % 2 -eq 0) {
                Set-ExcelRowColor -Worksheet $ws4 -Row $row -ColCount 9 -BgColor $Script:Colors.AltRowBg
            }
            $row++
        }
    }
    Set-ExcelColumnWidths -Worksheet $ws4 -Widths @(38, 28, 8, 16, 14, 14, 8, 22, 40)

    #─── 5. CONDITIONAL ACCESS ───────────────────────────────────────────────
    Write-Log "  Creazione: Conditional Access..." "INFO"
    $ws5 = New-Sheet "Conditional Access"

    Add-ExcelTitleRow -Worksheet $ws5 -Row 1 -Title "POLICY CONDITIONAL ACCESS" -MergeEnd 10

    $row = 3
    Add-ExcelHeaderRow -Worksheet $ws5 -Row $row `
        -Headers @("Nome Policy", "Stato", "Utenti Inclusi", "Applicazioni", "Piattaforme", "Tipi Client", "Controlli Grant", "Operatore", "Freq. Accesso", "Ultima Modifica")
    $row++

    if ($CaData) {
        foreach ($ca in ($CaData | Sort-Object State)) {
            $ws5.Cells[$row, 1].Value  = $ca.DisplayName
            $ws5.Cells[$row, 2].Value  = $ca.State
            $ws5.Cells[$row, 3].Value  = $ca.IncludeUsers
            $ws5.Cells[$row, 4].Value  = $ca.IncludeApplications
            $ws5.Cells[$row, 5].Value  = $ca.IncludePlatforms
            $ws5.Cells[$row, 6].Value  = $ca.ClientAppTypes
            $ws5.Cells[$row, 7].Value  = $ca.GrantControls
            $ws5.Cells[$row, 8].Value  = $ca.GrantOperator
            $ws5.Cells[$row, 9].Value  = $ca.SessionSignInFreq
            $ws5.Cells[$row, 10].Value = $ca.ModifiedDateTime

            $stateBg = switch ($ca.State) {
                "enabled"                              { $Script:Colors.OkBg }
                "enabledForReportingButNotEnforced"    { $Script:Colors.MediumBg }
                "disabled"                             { $Script:Colors.LowBg }
                default                                { $Script:Colors.InfoBg }
            }
            $stateFg = if ($ca.State -eq "enabled") { $Script:Colors.OkFg } else { "000000" }
            Set-ExcelCellColor -Worksheet $ws5 -Row $row -Col 2 -BgColor $stateBg -FgColor $stateFg
            $row++
        }
    }
    Set-ExcelColumnWidths -Worksheet $ws5 -Widths @(45, 14, 30, 30, 20, 28, 28, 12, 18, 22)

    #─── 6. LEGACY AUTHENTICATION ────────────────────────────────────────────
    Write-Log "  Creazione: Legacy Authentication..." "INFO"
    $ws6 = New-Sheet "Legacy Authentication"

    Add-ExcelTitleRow -Worksheet $ws6 -Row 1 -Title "AUTENTICAZIONE LEGACY - EVENTI RILEVATI" -MergeEnd 9

    $row = 3
    [object[]]$legacyArr = @()
    if ($null -ne $LegacyData) { $legacyArr = @($LegacyData) }
    if ($legacyArr.Count -gt 0) {
        $ws6.Cells[$row, 1, $row, 9].Merge = $true
        $ws6.Cells[$row, 1].Value = "ATTENZIONE: $($legacyArr.Count) eventi di autenticazione legacy rilevati negli ultimi $LookbackDays giorni"
        Set-ExcelCellColor -Worksheet $ws6 -Row $row -Col 1 -BgColor $Script:Colors.HighBg -FgColor "FFFFFF"
        $ws6.Cells[$row, 1].Style.Font.Bold = $true
        $row += 2

        Add-ExcelHeaderRow -Worksheet $ws6 -Row $row `
            -Headers @("UPN", "Display Name", "App", "Client Legacy", "Indirizzo IP", "Posizione", "Stato", "Livello Rischio", "Data/Ora")
        $row++

        foreach ($signin in ($legacyArr | Sort-Object CreatedDateTime -Descending)) {
            $ws6.Cells[$row, 1].Value = $signin.UPN
            $ws6.Cells[$row, 2].Value = $signin.DisplayName
            $ws6.Cells[$row, 3].Value = $signin.AppDisplayName
            $ws6.Cells[$row, 4].Value = $signin.ClientAppUsed
            $ws6.Cells[$row, 5].Value = $signin.IPAddress
            $ws6.Cells[$row, 6].Value = $signin.Location
            $ws6.Cells[$row, 7].Value = $signin.Status
            $ws6.Cells[$row, 8].Value = $signin.RiskLevel
            $ws6.Cells[$row, 9].Value = $signin.CreatedDateTime

            if ($signin.RiskLevel -in @("high", "medium")) {
                $bg = if ($signin.RiskLevel -eq "high") { $Script:Colors.HighBg } else { $Script:Colors.MediumBg }
                Set-ExcelRowColor -Worksheet $ws6 -Row $row -ColCount 9 -BgColor $bg -FgColor "FFFFFF"
            }
            $row++
        }
    } else {
        $ws6.Cells[$row, 1, $row, 9].Merge = $true
        $ws6.Cells[$row, 1].Value = "Nessun evento di autenticazione legacy rilevato nel periodo analizzato."
        Set-ExcelCellColor -Worksheet $ws6 -Row $row -Col 1 -BgColor $Script:Colors.OkBg -FgColor "FFFFFF"
        $ws6.Cells[$row, 1].Style.Font.Bold = $true
    }
    Set-ExcelColumnWidths -Worksheet $ws6 -Widths @(38, 25, 28, 26, 16, 22, 10, 14, 22)

    #─── 7. SMTP AUTH POP IMAP ───────────────────────────────────────────────
    Write-Log "  Creazione: SMTP AUTH POP IMAP..." "INFO"
    $ws7 = New-Sheet "SMTP AUTH POP IMAP"

    Add-ExcelTitleRow -Worksheet $ws7 -Row 1 -Title "CONFIGURAZIONE SMTP AUTH - POP3 - IMAP" -MergeEnd 9

    $row = 3
    $ws7.Cells[$row, 1, $row, 5].Merge = $true
    $ws7.Cells[$row, 1].Value = "CONFIGURAZIONE GLOBALE TENANT"
    $ws7.Cells[$row, 1].Style.Font.Bold = $true
    Set-ExcelCellColor -Worksheet $ws7 -Row $row -Col 1 `
        -BgColor $Script:Colors.SectionBg -FgColor $Script:Colors.SectionFg
    $row++

    $smtpGlobalDisabled = ($TransportData -and $TransportData.SmtpClientAuthenticationDisabled -eq $true)
    $ws7.Cells[$row, 1].Value = "SMTP AUTH Globale Disabilitato"
    $ws7.Cells[$row, 1].Style.Font.Bold = $true
    $ws7.Cells[$row, 2].Value = if ($smtpGlobalDisabled) { "SI - SICURO" } else { "NO - A RISCHIO" }
    $smtpBg = if ($smtpGlobalDisabled) { $Script:Colors.OkBg } else { $Script:Colors.CriticalBg }
    $smtpFg = if ($smtpGlobalDisabled) { $Script:Colors.OkFg } else { "FFFFFF" }
    Set-ExcelCellColor -Worksheet $ws7 -Row $row -Col 2 -BgColor $smtpBg -FgColor $smtpFg
    $row += 2

    [object[]]$casArr = @()
    if ($null -ne $CasData) { $casArr = @($CasData) }
    if ($casArr.Count -gt 0) {
        Add-ExcelHeaderRow -Worksheet $ws7 -Row $row `
            -Headers @("UPN", "Display Name", "SMTP Auth (Mbx)", "POP3", "IMAP", "ActiveSync", "OWA", "EWS", "MAPI")
        $row++

        $sortedCas = $casArr | Sort-Object {
            $s = 0
            if ($_.SmtpClientAuthenticationDisabled -eq $false) { $s += 3 }
            if ($_.PopEnabled  -eq $true) { $s += 2 }
            if ($_.ImapEnabled -eq $true) { $s += 2 }
            $s
        } -Descending

        foreach ($cas in $sortedCas) {
            $ws7.Cells[$row, 1].Value = $cas.UPN
            $ws7.Cells[$row, 2].Value = $cas.DisplayName
            $ws7.Cells[$row, 3].Value = if ($cas.SmtpClientAuthenticationDisabled -eq $true) { "Disabilitato" } `
                                        elseif ($cas.SmtpClientAuthenticationDisabled -eq $false) { "Abilitato" } `
                                        else { "Default (Tenant)" }
            $ws7.Cells[$row, 4].Value = if ($cas.PopEnabled)          { "Abilitato" } else { "Disabilitato" }
            $ws7.Cells[$row, 5].Value = if ($cas.ImapEnabled)         { "Abilitato" } else { "Disabilitato" }
            $ws7.Cells[$row, 6].Value = if ($cas.ActiveSyncEnabled)   { "Abilitato" } else { "Disabilitato" }
            $ws7.Cells[$row, 7].Value = if ($cas.OWAEnabled)          { "Abilitato" } else { "Disabilitato" }
            $ws7.Cells[$row, 8].Value = if ($cas.EwsEnabled)          { "Abilitato" } else { "Disabilitato" }
            $ws7.Cells[$row, 9].Value = if ($cas.MapiEnabled)         { "Abilitato" } else { "Disabilitato" }

            if ($cas.SmtpClientAuthenticationDisabled -eq $false) {
                Set-ExcelCellColor -Worksheet $ws7 -Row $row -Col 3 -BgColor $Script:Colors.HighBg -FgColor "FFFFFF"
            }
            if ($cas.PopEnabled -eq $true) {
                Set-ExcelCellColor -Worksheet $ws7 -Row $row -Col 4 -BgColor $Script:Colors.MediumBg
            }
            if ($cas.ImapEnabled -eq $true) {
                Set-ExcelCellColor -Worksheet $ws7 -Row $row -Col 5 -BgColor $Script:Colors.MediumBg
            }
            $row++
        }
    }
    Set-ExcelColumnWidths -Worksheet $ws7 -Widths @(38, 25, 18, 14, 14, 14, 12, 12, 12)

    #─── 8. GLOBAL ADMINISTRATORS ────────────────────────────────────────────
    Write-Log "  Creazione: Global Administrators..." "INFO"
    $ws8 = New-Sheet "Global Administrators"

    Add-ExcelTitleRow -Worksheet $ws8 -Row 1 -Title "GLOBAL ADMINISTRATORS E RUOLI PRIVILEGIATI" -MergeEnd 8

    $row = 3
    $ws8.Cells[$row, 1].Value = "Numero Global Administrators: $($ScoreResult.AdminCount)"
    $ws8.Cells[$row, 1].Style.Font.Bold = $true
    $adminBg = if ($ScoreResult.AdminCount -le 5) { $Script:Colors.OkBg } `
               elseif ($ScoreResult.AdminCount -le 10) { $Script:Colors.MediumBg } `
               else { $Script:Colors.CriticalBg }
    Set-ExcelCellColor -Worksheet $ws8 -Row $row -Col 1 -BgColor $adminBg -FgColor "FFFFFF"
    $row += 2

    if ($RolesData -and @($RolesData).Count -gt 0) {
        Add-ExcelHeaderRow -Worksheet $ws8 -Row $row `
            -Headers @("Ruolo", "UPN Membro", "Display Name", "Tipo Membro", "Account Attivo", "Tipo Utente", "Note Sicurezza")
        $row++

        $sortedRoles = @($RolesData) | Sort-Object {
            if ($_.RoleName -eq "Global Administrator") { 0 } else { 1 }
        }, RoleName

        foreach ($member in $sortedRoles) {
            $ws8.Cells[$row, 1].Value = $member.RoleName
            $ws8.Cells[$row, 2].Value = $member.MemberUPN
            $ws8.Cells[$row, 3].Value = $member.MemberDisplayName
            $ws8.Cells[$row, 4].Value = $member.MemberType
            $ws8.Cells[$row, 5].Value = $member.AccountEnabled
            $ws8.Cells[$row, 6].Value = $member.UserType

            $notes = [System.Collections.Generic.List[string]]::new()
            if ($member.AccountEnabled -eq $false) { $notes.Add("Account disabilitato") }
            if ($member.UserType -eq "Guest")       { $notes.Add("Utente guest con ruolo privilegiato") }
            if ($member.RoleName -eq "Global Administrator") { $notes.Add("Accesso completo al tenant") }
            $ws8.Cells[$row, 7].Value = ($notes -join "; ")

            if ($member.RoleName -eq "Global Administrator") {
                Set-ExcelRowColor -Worksheet $ws8 -Row $row -ColCount 7 `
                    -BgColor $Script:Colors.HighBg -FgColor "FFFFFF"
            }
            if ($member.UserType -eq "Guest") {
                Set-ExcelCellColor -Worksheet $ws8 -Row $row -Col 6 `
                    -BgColor $Script:Colors.CriticalBg -FgColor "FFFFFF"
            }
            $row++
        }
    }
    Set-ExcelColumnWidths -Worksheet $ws8 -Widths @(38, 35, 28, 28, 14, 14, 45)

    #─── 9. OAUTH GRANTS ─────────────────────────────────────────────────────
    Write-Log "  Creazione: OAuth Grants..." "INFO"
    $ws9 = New-Sheet "OAuth Grants"

    Add-ExcelTitleRow -Worksheet $ws9 -Row 1 -Title "GRANT OAUTH E SERVICE PRINCIPALS" -MergeEnd 7

    $row = 3
    [object[]]$oauthArr = @()
    if ($null -ne $OAuthData) { $oauthArr = @($OAuthData) }
    if ($oauthArr.Count -gt 0) {
        $highRiskCount = @($oauthArr | Where-Object { $_.RiskLevel -eq "ALTO" }).Count
        $ws9.Cells[$row, 1].Value = "Grant totali: $($oauthArr.Count)  |  Ad alto rischio: $highRiskCount"
        $ws9.Cells[$row, 1].Style.Font.Bold = $true
        if ($highRiskCount -gt 0) {
            Set-ExcelCellColor -Worksheet $ws9 -Row $row -Col 1 -BgColor $Script:Colors.HighBg -FgColor "FFFFFF"
        }
        $row += 2

        Add-ExcelHeaderRow -Worksheet $ws9 -Row $row `
            -Headers @("App (Client)", "Publisher", "Tipo Consenso", "Scope Assegnati", "Scope Alto Rischio", "Livello Rischio", "Scadenza")
        $row++

        foreach ($grant in ($oauthArr | Sort-Object RiskLevel -Descending)) {
            $ws9.Cells[$row, 1].Value = $grant.ClientDisplayName
            $ws9.Cells[$row, 2].Value = $grant.ClientPublisher
            $ws9.Cells[$row, 3].Value = $grant.ConsentType
            $ws9.Cells[$row, 4].Value = $grant.Scopes
            $ws9.Cells[$row, 5].Value = $grant.HighRiskScopes
            $ws9.Cells[$row, 6].Value = $grant.RiskLevel
            $ws9.Cells[$row, 7].Value = $grant.ExpiryTime

            if ($grant.RiskLevel -eq "ALTO") {
                Set-ExcelRowColor -Worksheet $ws9 -Row $row -ColCount 7 `
                    -BgColor $Script:Colors.HighBg -FgColor "FFFFFF"
            } elseif ($row % 2 -eq 0) {
                Set-ExcelRowColor -Worksheet $ws9 -Row $row -ColCount 7 -BgColor $Script:Colors.AltRowBg
            }
            $row++
        }
    }
    Set-ExcelColumnWidths -Worksheet $ws9 -Widths @(38, 30, 16, 70, 50, 14, 22)

    #─── 10. SIGN-IN ANOMALIES ───────────────────────────────────────────────
    Write-Log "  Creazione: Sign-in Anomalies..." "INFO"
    $ws10 = New-Sheet "Sign-in Anomalies"

    Add-ExcelTitleRow -Worksheet $ws10 -Row 1 -Title "ANOMALIE SIGN-IN E ACCESSI A RISCHIO" -MergeEnd 9

    $row = 3
    [object[]]$signInArr = @()
    if ($null -ne $SignInData) { $signInArr = @($SignInData) }
    if ($signInArr.Count -gt 0) {
        $anomalies = @($signInArr | Where-Object {
            ($_.RiskLevel -notin @("none", "", $null)) -or
            ($_.ConditionalAccessStatus -eq "failure") -or
            ($_.Status -notin @(0, "0", "", $null))
        })

        $ws10.Cells[$row, 1].Value = "Log totali analizzati: $($signInArr.Count)  |  Anomalie rilevate: $($anomalies.Count)"
        $ws10.Cells[$row, 1].Style.Font.Bold = $true
        if ($anomalies.Count -gt 0) {
            Set-ExcelCellColor -Worksheet $ws10 -Row $row -Col 1 -BgColor $Script:Colors.MediumBg
        }
        $row += 2

        Add-ExcelHeaderRow -Worksheet $ws10 -Row $row `
            -Headers @("UPN", "App", "Client", "Indirizzo IP", "Posizione", "Stato", "Livello Rischio", "Stato CA", "Data/Ora")
        $row++

        $displayData = if ($anomalies.Count -gt 0) { $anomalies | Select-Object -First 500 } else { $signInArr | Select-Object -First 200 }
        foreach ($signin in ($displayData | Sort-Object CreatedDateTime -Descending)) {
            $ws10.Cells[$row, 1].Value = $signin.UPN
            $ws10.Cells[$row, 2].Value = $signin.AppDisplayName
            $ws10.Cells[$row, 3].Value = $signin.ClientAppUsed
            $ws10.Cells[$row, 4].Value = $signin.IPAddress
            $ws10.Cells[$row, 5].Value = $signin.Location
            $ws10.Cells[$row, 6].Value = $signin.Status
            $ws10.Cells[$row, 7].Value = $signin.RiskLevel
            $ws10.Cells[$row, 8].Value = $signin.ConditionalAccessStatus
            $ws10.Cells[$row, 9].Value = $signin.CreatedDateTime

            if ($signin.RiskLevel -eq "high") {
                Set-ExcelRowColor -Worksheet $ws10 -Row $row -ColCount 9 `
                    -BgColor $Script:Colors.CriticalBg -FgColor "FFFFFF"
            } elseif ($signin.RiskLevel -eq "medium") {
                Set-ExcelRowColor -Worksheet $ws10 -Row $row -ColCount 9 -BgColor $Script:Colors.MediumBg
            }
            $row++
        }
    }
    Set-ExcelColumnWidths -Worksheet $ws10 -Widths @(38, 28, 22, 16, 24, 10, 14, 18, 22)

    #─── 11. MAILBOX RULES FORWARDING ────────────────────────────────────────
    Write-Log "  Creazione: Mailbox Rules Forwarding..." "INFO"
    $ws11 = New-Sheet "Mailbox Rules Forwarding"

    Add-ExcelTitleRow -Worksheet $ws11 -Row 1 -Title "REGOLE INBOX E FORWARDING SOSPETTI" -MergeEnd 8

    $row = 3
    [object[]]$fwdArr2 = @()
    if ($RulesData -and $RulesData.ContainsKey("Forwarding") -and $null -ne $RulesData["Forwarding"]) {
        $fwdArr2 = @($RulesData["Forwarding"])
    }
    [object[]]$rulesArr2 = @()
    if ($RulesData -and $RulesData.ContainsKey("Rules") -and $null -ne $RulesData["Rules"]) {
        $rulesArr2 = @($RulesData["Rules"])
    }

    if ($fwdArr2.Count -gt 0) {
        $ws11.Cells[$row, 1, $row, 8].Merge = $true
        $ws11.Cells[$row, 1].Value = "FORWARDING ESTERNO ATTIVO: $($fwdArr2.Count) mailbox con forward configurato"
        Set-ExcelCellColor -Worksheet $ws11 -Row $row -Col 1 -BgColor $Script:Colors.HighBg -FgColor "FFFFFF"
        $ws11.Cells[$row, 1].Style.Font.Bold = $true
        $row += 2

        Add-ExcelHeaderRow -Worksheet $ws11 -Row $row `
            -Headers @("UPN", "Display Name", "Forwarding Address", "SMTP Forward", "Mantieni Copia", "Fonte", "Livello Rischio")
        $row++
        foreach ($fwd in $fwdArr2) {
            $ws11.Cells[$row, 1].Value = $fwd.UPN
            $ws11.Cells[$row, 2].Value = $fwd.DisplayName
            $ws11.Cells[$row, 3].Value = $fwd.ForwardingAddress
            $ws11.Cells[$row, 4].Value = $fwd.ForwardingSmtpAddress
            $ws11.Cells[$row, 5].Value = $fwd.DeliverToMailboxAndForward
            $ws11.Cells[$row, 6].Value = $fwd.Source
            $ws11.Cells[$row, 7].Value = $fwd.RiskLevel
            $rcp = Get-RiskColors -Level $fwd.RiskLevel
            Set-ExcelRowColor -Worksheet $ws11 -Row $row -ColCount 7 -BgColor $rcp[0] -FgColor $rcp[1]
            $row++
        }
        $row += 2
    }

    $riskyRules = @($rulesArr2 | Where-Object { $_.RiskLevel -eq "ALTO" })
    if ($riskyRules.Count -gt 0) {
        $ws11.Cells[$row, 1, $row, 8].Merge = $true
        $ws11.Cells[$row, 1].Value = "REGOLE INBOX SOSPETTE: $($riskyRules.Count) regole con azioni potenzialmente pericolose"
        Set-ExcelCellColor -Worksheet $ws11 -Row $row -Col 1 -BgColor $Script:Colors.MediumBg
        $ws11.Cells[$row, 1].Style.Font.Bold = $true
        $row += 2

        Add-ExcelHeaderRow -Worksheet $ws11 -Row $row `
            -Headers @("UPN", "Nome Regola", "Attiva", "Forward A", "Reindirizza A", "Elimina", "Parole Chiave", "Rischio")
        $row++
        foreach ($rule in $riskyRules) {
            $ws11.Cells[$row, 1].Value = $rule.UPN
            $ws11.Cells[$row, 2].Value = $rule.RuleName
            $ws11.Cells[$row, 3].Value = $rule.RuleEnabled
            $ws11.Cells[$row, 4].Value = $rule.ForwardTo
            $ws11.Cells[$row, 5].Value = $rule.RedirectTo
            $ws11.Cells[$row, 6].Value = $rule.DeleteMessage
            $ws11.Cells[$row, 7].Value = $rule.SubjectContainsWords
            $ws11.Cells[$row, 8].Value = $rule.RiskLevel
            Set-ExcelRowColor -Worksheet $ws11 -Row $row -ColCount 8 `
                -BgColor $Script:Colors.HighBg -FgColor "FFFFFF"
            $row++
        }
    } elseif ($fwdArr2.Count -eq 0) {
        $ws11.Cells[$row, 1, $row, 8].Merge = $true
        $ws11.Cells[$row, 1].Value = "Nessuna regola inbox sospetta o forwarding rilevato."
        Set-ExcelCellColor -Worksheet $ws11 -Row $row -Col 1 -BgColor $Script:Colors.OkBg -FgColor "FFFFFF"
    }
    Set-ExcelColumnWidths -Worksheet $ws11 -Widths @(38, 30, 8, 35, 35, 10, 28, 12)

    #─── 12. RISK DETECTIONS ─────────────────────────────────────────────────
    Write-Log "  Creazione: Risk Detections..." "INFO"
    $ws12 = New-Sheet "Risk Detections"

    Add-ExcelTitleRow -Worksheet $ws12 -Row 1 -Title "RILEVAMENTI RISCHI IDENTITA (Identity Protection)" -MergeEnd 9

    $row = 3
    [object[]]$riskDetArr = @()
    if ($null -ne $RiskDetectionsData) { $riskDetArr = @($RiskDetectionsData) }
    if ($riskDetArr.Count -gt 0) {
        $ws12.Cells[$row, 1].Value = "Rilevamenti totali nel periodo: $($riskDetArr.Count)"
        $ws12.Cells[$row, 1].Style.Font.Bold = $true
        Set-ExcelCellColor -Worksheet $ws12 -Row $row -Col 1 -BgColor $Script:Colors.HighBg -FgColor "FFFFFF"
        $row += 2

        Add-ExcelHeaderRow -Worksheet $ws12 -Row $row `
            -Headers @("UPN", "Tipo Rischio", "Livello", "Stato", "Dettaglio", "Indirizzo IP", "Posizione", "Fonte", "Data Rilevamento")
        $row++

        foreach ($det in ($riskDetArr | Sort-Object RiskLevel, DetectedDateTime -Descending)) {
            $ws12.Cells[$row, 1].Value = $det.UPN
            $ws12.Cells[$row, 2].Value = $det.RiskType
            $ws12.Cells[$row, 3].Value = $det.RiskLevel
            $ws12.Cells[$row, 4].Value = $det.RiskState
            $ws12.Cells[$row, 5].Value = $det.RiskDetail
            $ws12.Cells[$row, 6].Value = $det.IPAddress
            $ws12.Cells[$row, 7].Value = $det.Location
            $ws12.Cells[$row, 8].Value = $det.Source
            $ws12.Cells[$row, 9].Value = $det.DetectedDateTime

            $riskBg = switch ($det.RiskLevel) {
                "high"   { $Script:Colors.CriticalBg }
                "medium" { $Script:Colors.MediumBg }
                default  { $Script:Colors.InfoBg }
            }
            $riskFg = if ($det.RiskLevel -eq "high") { "FFFFFF" } else { "000000" }
            Set-ExcelRowColor -Worksheet $ws12 -Row $row -ColCount 9 -BgColor $riskBg -FgColor $riskFg
            $row++
        }
    } else {
        $ws12.Cells[$row, 1, $row, 9].Merge = $true
        $ws12.Cells[$row, 1].Value = "Nessun rilevamento rischi nel periodo (o licenza P2 non disponibile)."
        Set-ExcelCellColor -Worksheet $ws12 -Row $row -Col 1 -BgColor $Script:Colors.OkBg -FgColor "FFFFFF"
    }
    Set-ExcelColumnWidths -Worksheet $ws12 -Widths @(38, 32, 12, 14, 28, 16, 22, 16, 22)

    #─── 13. RAW DATA (opzionale) ────────────────────────────────────────────
    if ($IncludeRawData) {
        Write-Log "  Creazione: Raw Data (richiesto)..." "INFO"
        $ws13 = New-Sheet "Raw Data"

        Add-ExcelTitleRow -Worksheet $ws13 -Row 1 -Title "DATI GREZZI - PER ANALISI TECNICA AVANZATA" -MergeEnd 4

        $row = 3
        $ws13.Cells[$row, 1].Value = "Statistiche Raccolta Dati"
        $ws13.Cells[$row, 1].Style.Font.Bold = $true
        $row++

        $rulesCountRaw = 0
        if ($RulesData -and $RulesData.ContainsKey("Rules") -and $null -ne $RulesData["Rules"]) {
            $rulesCountRaw = @($RulesData["Rules"]).Count
        }
        $fwdCountRaw = 0
        if ($RulesData -and $RulesData.ContainsKey("Forwarding") -and $null -ne $RulesData["Forwarding"]) {
            $fwdCountRaw = @($RulesData["Forwarding"]).Count
        }

        $mfaDataCount            = if ($MfaData) { @($MfaData).Count } else { 0 }
        $caDataCount             = if ($CaData) { @($CaData).Count } else { 0 }
        $signInDataCount         = if ($SignInData) { @($SignInData).Count } else { 0 }
        $legacyDataCount         = if ($LegacyData) { @($LegacyData).Count } else { 0 }
        $casDataCount            = if ($CasData) { @($CasData).Count } else { 0 }
        $mailboxDataCount        = if ($MailboxData) { @($MailboxData).Count } else { 0 }
        $rolesDataCount          = if ($RolesData) { @($RolesData).Count } else { 0 }
        $oauthDataCount          = if ($OAuthData) { @($OAuthData).Count } else { 0 }
        $riskyUsersDataCount     = if ($RiskyUsersData) { @($RiskyUsersData).Count } else { 0 }
        $riskDetectionsDataCount = if ($RiskDetectionsData) { @($RiskDetectionsData).Count } else { 0 }

        $rawStats = @(
            @("Utenti MFA analizzati",  $mfaDataCount),
            @("Policy CA",             $caDataCount),
            @("Sign-in logs",          $signInDataCount),
            @("Eventi legacy auth",    $legacyDataCount),
            @("CAS Mailbox",           $casDataCount),
            @("Mailbox totali",        $mailboxDataCount),
            @("Role assignments",      $rolesDataCount),
            @("OAuth Grants",          $oauthDataCount),
            @("Risky Users",           $riskyUsersDataCount),
            @("Risk Detections",       $riskDetectionsDataCount),
            @("Inbox Rules",           $rulesCountRaw),
            @("Forward configurati",   $fwdCountRaw)
        )

        foreach ($stat in $rawStats) {
            $ws13.Cells[$row, 1].Value = $stat[0]
            $ws13.Cells[$row, 2].Value = $stat[1]
            $row++
        }
        Set-ExcelColumnWidths -Worksheet $ws13 -Widths @(40, 15, 40, 15)
    }

    # Imposta il primo foglio come attivo
    $excelPkg.Workbook.Worksheets["Sintesi Esecutiva"].Select()

    # Chiude e salva tramite Close-ExcelPackage di ImportExcel (gestisce Dispose internamente)
    Close-ExcelPackage $excelPkg

    Write-Log "Workbook Excel salvato: $ExcelPath" "SUCCESS"
    return $ExcelPath
}

#endregion

#region ─── REPORT ESECUTIVO MD ───────────────────────────────────────────────

function New-ExecutiveReport {
    param(
        [string]$ReportPath,
        $ScoreResult,
        $MfaData,
        $CaData,
        $TransportData,
        $RolesData,
        $LegacyData,
        $OAuthData,
        $RulesData
    )

    $duration      = (Get-Date) - $Script:StartTime
    $legacyCount = 0
    if ($null -ne $LegacyData) { $legacyCount = @($LegacyData).Count }
    $highRiskOAuth = 0
    if ($null -ne $OAuthData) { $highRiskOAuth = @($OAuthData | Where-Object { $_.RiskLevel -eq "ALTO" }).Count }
    $riskyFwdCount = 0
    if ($RulesData -and $RulesData.ContainsKey("Forwarding") -and $null -ne $RulesData["Forwarding"]) {
        $riskyFwdCount = @($RulesData["Forwarding"]).Count
    }
    $enabledCa = 0
    if ($null -ne $CaData) { $enabledCa = @($CaData | Where-Object { $_.State -eq "enabled" }).Count }
    $totalCa = 0
    if ($null -ne $CaData) { $totalCa = @($CaData).Count }
    $smtpStatus    = if ($TransportData -and $TransportData.SmtpClientAuthenticationDisabled -eq $true) { "DISABILITATO - Configurazione sicura" } else { "ABILITATO - Rischio autenticazione legacy e credential spray" }
    $mfaCovMsg     = if ($ScoreResult.MfaStats.Coverage -ge 90) { "BUONO - Alta copertura MFA" } elseif ($ScoreResult.MfaStats.Coverage -ge 70) { "MEDIO - Copertura sufficiente ma migliorabile" } else { "CRITICO - Copertura MFA insufficiente" }
    $adminMsg      = if ($ScoreResult.AdminCount -le 5 -and $ScoreResult.AdminCount -ge 2) { "ADEGUATO (2-5 admin)" } elseif ($ScoreResult.AdminCount -gt 5) { "ELEVATO - Ridurre il numero di Global Admin" } else { "CRITICO - Numero di admin non ottimale" }
    $caMsg         = if ($enabledCa -gt 0) { "Policy CA configurate" } else { "CRITICO - Nessuna policy CA attiva" }
    $legacyMsg     = if ($legacyCount -eq 0) { "Nessun evento legacy rilevato" } else { "$legacyCount eventi di autenticazione legacy - Valutare blocco" }
    $oauthMsg      = if ($highRiskOAuth -eq 0) { "Nessun grant OAuth ad alto rischio" } else { "$highRiskOAuth grant OAuth con scope pericolosi - Revisione necessaria" }
    $fwdMsg        = if ($riskyFwdCount -eq 0) { "Nessun forwarding sospetto" } else { "$riskyFwdCount regole di forwarding attive - Verificare autorizzazione" }
    $durMinutes    = [math]::Round($duration.TotalMinutes, 1)
    $reportDate    = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    $incidentLine  = if (-not [string]::IsNullOrEmpty($IncidentUser)) { $IncidentUser } else { "N/A (assessment generale)" }

    $lines = @(
        "================================================================================"
        "  VALUTAZIONE SICUREZZA MICROSOFT 365"
        "  RAPPORTO ESECUTIVO RISERVATO"
        "================================================================================"
        ""
        "INFORMAZIONI ASSESSMENT"
        "--------------------------------------------------------------------------------"
        "Tenant analizzato  : $TenantName"
        "Data assessment    : $(Get-Date -Format 'dd MMMM yyyy alle HH:mm')"
        "Periodo analisi    : Ultimi $LookbackDays giorni"
        "Durata assessment  : $durMinutes minuti"
        "Versione script    : v$($Script:AssessmentVersion)"
        "Utente indagine    : $incidentLine"
        ""
        "================================================================================"
        "SECURITY SCORE: $($ScoreResult.Score)/$($ScoreResult.MaxScore) ($($ScoreResult.Percentage)%) - RATING: $($ScoreResult.Rating)"
        "================================================================================"
        ""
        "SINTESI RISCHI PRINCIPALI"
        "--------------------------------------------------------------------------------"
        ""
        "1. AUTENTICAZIONE MULTI-FATTORE (MFA)"
        "   Copertura: $($ScoreResult.MfaStats.Coverage)% ($($ScoreResult.MfaStats.Registered)/$($ScoreResult.MfaStats.Total) utenti)"
        "   Stato: $mfaCovMsg"
        ""
        "2. SMTP AUTHENTICATION (Legacy)"
        "   Stato globale: $smtpStatus"
        ""
        "3. GLOBAL ADMINISTRATORS"
        "   Numero: $($ScoreResult.AdminCount) Global Administrator rilevati"
        "   Stato: $adminMsg"
        ""
        "4. CONDITIONAL ACCESS"
        "   Policy attive: $enabledCa / $totalCa totali"
        "   Stato: $caMsg"
        ""
        "5. AUTENTICAZIONE LEGACY"
        "   Eventi rilevati: $legacyCount negli ultimi $LookbackDays giorni"
        "   Stato: $legacyMsg"
        ""
        "6. OAUTH GRANTS AD ALTO RISCHIO"
        "   Grant a rischio: $highRiskOAuth"
        "   Stato: $oauthMsg"
        ""
        "7. FORWARDING & REGOLE INBOX"
        "   Forward attivi: $riskyFwdCount"
        "   Stato: $fwdMsg"
        ""
        "================================================================================"
        "AZIONI PRIORITARIE RACCOMANDATE"
        "================================================================================"
        ""
        "PRIORITA 1 - CRITICA (Immediato / entro 7 giorni):"
        "  - Abilitare MFA per tutti gli utenti non registrati"
        "  - Creare policy Conditional Access per richiedere MFA globalmente"
        "  - Verificare e revocare grant OAuth con scope ad alto rischio"
        ""
        "PRIORITA 2 - ALTA (entro 30 giorni):"
        "  - Disabilitare SMTP AUTH a livello tenant (se non gia fatto)"
        "  - Creare policy CA per bloccare autenticazione legacy"
        "  - Ridurre il numero di Global Administrator"
        "  - Verificare regole di forwarding e inbox rules sospette"
        ""
        "PRIORITA 3 - MEDIA (entro 60 giorni):"
        "  - Disabilitare POP3/IMAP per mailbox che non ne necessitano"
        "  - Implementare Microsoft Defender for Office 365"
        "  - Configurare alert di sicurezza per eventi critici"
        ""
        "PRIORITA 4 - BASSA (entro 90 giorni):"
        "  - Valutare implementazione PIM per ruoli privilegiati"
        "  - Abilitare Continuous Access Evaluation (CAE)"
        "  - Configurare integrazione SIEM"
        ""
        "================================================================================"
        "DICHIARAZIONE DI LIMITAZIONE"
        "================================================================================"
        ""
        "Questo report e stato generato da uno script automatizzato read-only."
        "I dati raccolti riflettono lo stato del tenant al momento dell'assessment."
        "Le raccomandazioni devono essere validate da un esperto di sicurezza prima"
        "dell'implementazione. Nessuna modifica e stata apportata al tenant."
        ""
        "Report generato il: $reportDate"
        "================================================================================"
    )

    $lines | Set-Content -Path $ReportPath -Encoding UTF8
    Write-Log "Report esecutivo salvato: $ReportPath" "SUCCESS"
    return $ReportPath
}

#endregion

#region ─── HTML REPORT ESECUTIVO ─────────────────────────────────────────────

function New-HtmlReport {
    <#
    .SYNOPSIS
        Genera un report HTML esecutivo in italiano per il security assessment M365.
    .DESCRIPTION
        Report visuale con score, rating, rischi principali e azioni prioritarie.
        Non include dati sensibili estesi. Solo KPI e sommari aggregati.
    #>
    param(
        [string]$HtmlPath,
        $ScoreResult,
        $MfaData,
        $CaData,
        $TransportData,
        $RolesData,
        $LegacyData,
        $OAuthData,
        $RulesData,
        $CasData
    )

    Write-Progress -Activity "Generazione HTML Report" -Status "Calcolo metriche..." -PercentComplete 10

    # ── Calcolo metriche ───────────────────────────────────────────────────────
    if ($LegacyData) {
        $legacyCount = @($LegacyData).Count
    } else {
        $legacyCount = 0
    }
    if ($OAuthData) {
        $highRiskOAuth = @($OAuthData | Where-Object { $_.RiskLevel -eq "ALTO" }).Count
    } else {
        $highRiskOAuth = 0
    }
    if ($OAuthData) {
        $totalOAuth = @($OAuthData).Count
    } else {
        $totalOAuth = 0
    }
    if ($RulesData -and $RulesData.ContainsKey("Forwarding") -and $RulesData["Forwarding"]) {
        $riskyFwdCount = @($RulesData["Forwarding"]).Count
    } else {
        $riskyFwdCount = 0
    }
    if ($CaData) {
        $enabledCa = @($CaData | Where-Object { $_.State -eq "enabled" }).Count
    } else {
        $enabledCa = 0
    }
    if ($CaData) {
        $totalCa = @($CaData).Count
    } else {
        $totalCa = 0
    }
    $smtpOk = ($TransportData -and $TransportData.SmtpClientAuthenticationDisabled -eq $true)
    if ($CasData) {
        $popImapCount = @($CasData | Where-Object { $_.PopEnabled -eq $true -or $_.ImapEnabled -eq $true }).Count
    } else {
        $popImapCount = 0
    }
    if ($CasData) {
        $casTotal = @($CasData).Count
    } else {
        $casTotal = 0
    }
    if ($MfaData) {
        $adminSenzaMfa = @($MfaData | Where-Object { $_.IsAdmin -eq $true -and $_.IsMfaRegistered -ne $true }).Count
    } else {
        $adminSenzaMfa = 0
    }

    $dateStr    = Get-Date -Format "dd/MM/yyyy HH:mm"
    $dateFile   = Get-Date -Format "dd MMMM yyyy"
    $pct        = $ScoreResult.Percentage
    $rating     = $ScoreResult.Rating
    $score      = $ScoreResult.Score
    $maxScore   = $ScoreResult.MaxScore
    $mfaCov     = $ScoreResult.MfaStats.Coverage
    $adminCount = $ScoreResult.AdminCount

    # ── Colori del rating ─────────────────────────────────────────────────────
    $ratingColor = switch ($rating) {
        "BUONO"        { "#1a7a1a" }
        "SUFFICIENTE"  { "#b38a00" }
        "INSUFFICIENTE"{ "#cc4400" }
        default        { "#a00000" }
    }
    $ratingBg = switch ($rating) {
        "BUONO"        { "#d4edda" }
        "SUFFICIENTE"  { "#fff3cd" }
        "INSUFFICIENTE"{ "#ffe5d0" }
        default        { "#f8d7da" }
    }

    # ── Helper: row per tabella rischi ────────────────────────────────────────
    function New-RiskRow {
        param([string]$Area, [string]$Status, [string]$Level, [string]$Detail)
        $bg = switch ($Level) {
            "CRITICO" { "#f8d7da" } "ALTO" { "#fff0d0" } "MEDIO" { "#fff3cd" }
            "OK"      { "#d4edda" } default { "#e8f0fe" }
        }
        $badge = switch ($Level) {
            "CRITICO" { "#a00000" } "ALTO" { "#cc4400" } "MEDIO" { "#b38a00" }
            "OK"      { "#1a7a1a" } default { "#2c5282" }
        }
        return @"
<tr style="background:$bg">
  <td style="padding:8px 12px;font-weight:600">$Area</td>
  <td style="padding:8px 12px">$Status</td>
  <td style="padding:8px 12px;text-align:center">
    <span style="background:$badge;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.78em;font-weight:700">$Level</span>
  </td>
  <td style="padding:8px 12px;color:#444;font-size:0.9em">$Detail</td>
</tr>
"@
    }

    Write-Progress -Activity "Generazione HTML Report" -Status "Composizione sezioni rischi..." -PercentComplete 40

    # ── Righe rischi ──────────────────────────────────────────────────────────
    if ($mfaCov -ge 90) {
        $mfaLevel = "OK"
    } elseif ($mfaCov -ge 70) {
        $mfaLevel = "MEDIO"
    } else {
        $mfaLevel = "CRITICO"
    }
    if ($enabledCa -gt 0) {
        $caLevel = "OK"
    } else {
        $caLevel = "CRITICO"
    }
    if ($smtpOk) {
        $smtpLevel = "OK"
    } else {
        $smtpLevel = "ALTO"
    }
    if ($highRiskOAuth -eq 0) {
        $oauthLevel = "OK"
    } elseif ($highRiskOAuth -le 3) {
        $oauthLevel = "MEDIO"
    } else {
        $oauthLevel = "ALTO"
    }
    if ($adminCount -le 5 -and $adminCount -ge 2) {
        $adminLevel = "OK"
    } elseif ($adminCount -gt 10) {
        $adminLevel = "ALTO"
    } else {
        $adminLevel = "MEDIO"
    }
    if ($legacyCount -eq 0) {
        $legacyLevel = "OK"
    } elseif ($legacyCount -lt 10) {
        $legacyLevel = "MEDIO"
    } else {
        $legacyLevel = "ALTO"
    }
    if ($riskyFwdCount -eq 0) {
        $fwdLevel = "OK"
    } else {
        $fwdLevel = "ALTO"
    }
    if ($casTotal -eq 0 -or ($popImapCount / [math]::Max($casTotal, 1) * 100) -lt 5) {
        $popLevel = "OK"
    } else {
        $popLevel = "MEDIO"
    }

    if ($smtpOk) {
        $smtpDetail = "Disabilitato - configurazione sicura"
        $smtpAuthStatus = "OFF"
    } else {
        $smtpDetail = "Abilitato - rischio credential spray"
        $smtpAuthStatus = "ON"
    }

    $riskRows = (
        (New-RiskRow "Autenticazione MFA" "$mfaCov% copertura ($($ScoreResult.MfaStats.Registered)/$($ScoreResult.MfaStats.Total) utenti)" $mfaLevel "Admin senza MFA: $adminSenzaMfa") +
        (New-RiskRow "Conditional Access" "$enabledCa policy attive su $totalCa totali" $caLevel "Verificare policy per MFA e legacy auth") +
        (New-RiskRow "SMTP Authentication" $smtpDetail $smtpLevel "") +
        (New-RiskRow "OAuth Grants" "$highRiskOAuth ad alto rischio su $totalOAuth totali" $oauthLevel "Scope pericolosi: Mail.ReadWrite, Files.ReadWrite.All") +
        (New-RiskRow "Global Administrators" "$adminCount Global Admin rilevati" $adminLevel "Ottimale: 2-5 amministratori") +
        (New-RiskRow "Autenticazione Legacy" "$legacyCount eventi negli ultimi $LookbackDays giorni" $legacyLevel "") +
        (New-RiskRow "Forwarding & Inbox Rules" "$riskyFwdCount forward attivi" $fwdLevel "Verificare autorizzazione") +
        (New-RiskRow "POP3 / IMAP" "$popImapCount mailbox esposte su $casTotal" $popLevel "Disabilitare per mailbox non legacy")
    )

    # ── Score gauge (barra orizzontale CSS) ───────────────────────────────────
    $gaugeColor = switch ($rating) {
        "BUONO"        { "#1a7a1a" }
        "SUFFICIENTE"  { "#b38a00" }
        "INSUFFICIENTE"{ "#cc4400" }
        default        { "#a00000" }
    }

    # ── Righe azioni prioritarie ──────────────────────────────────────────────
    $actionRows = @"
<tr style="background:#f8d7da"><td style="padding:8px 12px"><span style="background:#a00000;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.78em;font-weight:700">1 - CRITICA</span></td><td style="padding:8px 12px">MFA / CA Policy</td><td style="padding:8px 12px">Abilitare MFA per tutti gli utenti e creare policy CA obbligatoria</td><td style="padding:8px 12px;text-align:center">Immediato</td></tr>
<tr style="background:#fff0d0"><td style="padding:8px 12px"><span style="background:#cc4400;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.78em;font-weight:700">2 - ALTA</span></td><td style="padding:8px 12px">Legacy Auth</td><td style="padding:8px 12px">Bloccare autenticazione legacy tramite Conditional Access</td><td style="padding:8px 12px;text-align:center">14 giorni</td></tr>
<tr style="background:#fff0d0"><td style="padding:8px 12px"><span style="background:#cc4400;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.78em;font-weight:700">2 - ALTA</span></td><td style="padding:8px 12px">SMTP AUTH</td><td style="padding:8px 12px">Disabilitare SMTP AUTH globalmente in Exchange Online</td><td style="padding:8px 12px;text-align:center">30 giorni</td></tr>
<tr style="background:#fff0d0"><td style="padding:8px 12px"><span style="background:#cc4400;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.78em;font-weight:700">2 - ALTA</span></td><td style="padding:8px 12px">OAuth</td><td style="padding:8px 12px">Revisionare e revocare grant OAuth con scope ad alto rischio</td><td style="padding:8px 12px;text-align:center">30 giorni</td></tr>
<tr style="background:#fff3cd"><td style="padding:8px 12px"><span style="background:#b38a00;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.78em;font-weight:700">3 - MEDIA</span></td><td style="padding:8px 12px">Admin</td><td style="padding:8px 12px">Ridurre Global Admin, usare ruoli specifici e valutare PIM</td><td style="padding:8px 12px;text-align:center">60 giorni</td></tr>
<tr style="background:#fff3cd"><td style="padding:8px 12px"><span style="background:#b38a00;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.78em;font-weight:700">3 - MEDIA</span></td><td style="padding:8px 12px">POP3 / IMAP</td><td style="padding:8px 12px">Disabilitare POP3 e IMAP per mailbox che non ne necessitano</td><td style="padding:8px 12px;text-align:center">60 giorni</td></tr>
<tr style="background:#d4edda"><td style="padding:8px 12px"><span style="background:#1a7a1a;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.78em;font-weight:700">4 - BASSA</span></td><td style="padding:8px 12px">Monitoring</td><td style="padding:8px 12px">Configurare Microsoft Defender for Office 365 e alert SIEM</td><td style="padding:8px 12px;text-align:center">90 giorni</td></tr>
"@

    Write-Progress -Activity "Generazione HTML Report" -Status "Scrittura file HTML..." -PercentComplete 70

    # ── Template HTML ─────────────────────────────────────────────────────────
    $html = @"
<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>M365 Security Assessment - $TenantName</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Segoe UI',Arial,sans-serif;background:#f0f4f8;color:#1a202c;font-size:15px}
  .header{background:linear-gradient(135deg,#1F4E79 0%,#2E75B6 100%);color:#fff;padding:32px 40px}
  .header h1{font-size:1.6em;font-weight:700;letter-spacing:.5px}
  .header .sub{font-size:0.92em;margin-top:6px;opacity:.85}
  .container{max-width:1100px;margin:0 auto;padding:28px 24px}
  .card{background:#fff;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.07);margin-bottom:24px;overflow:hidden}
  .card-header{background:#1F4E79;color:#fff;padding:14px 20px;font-weight:700;font-size:1em;letter-spacing:.3px}
  .card-body{padding:20px}
  .score-box{display:flex;align-items:center;gap:32px;flex-wrap:wrap}
  .score-circle{width:120px;height:120px;border-radius:50%;display:flex;flex-direction:column;align-items:center;justify-content:center;background:$ratingBg;border:5px solid $ratingColor;flex-shrink:0}
  .score-num{font-size:2em;font-weight:800;color:$ratingColor}
  .score-label{font-size:0.7em;color:$ratingColor;font-weight:600;margin-top:2px}
  .score-details{flex:1}
  .score-details h2{font-size:1.5em;font-weight:700;color:$ratingColor}
  .score-details .sub{color:#555;font-size:0.9em;margin-top:4px}
  .gauge-wrap{margin-top:14px;background:#e2e8f0;border-radius:8px;height:16px;overflow:hidden}
  .gauge-fill{height:100%;background:$gaugeColor;border-radius:8px;width:$pct%;transition:width .8s}
  .gauge-lbl{font-size:0.78em;color:#666;margin-top:4px}
  .kpi-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:16px}
  .kpi{background:#f8fafc;border-radius:8px;padding:16px;border-left:4px solid #2E75B6;text-align:center}
  .kpi .val{font-size:2em;font-weight:800;color:#1F4E79}
  .kpi .lbl{font-size:0.78em;color:#555;margin-top:4px}
  table{width:100%;border-collapse:collapse;font-size:0.88em}
  th{background:#1F4E79;color:#fff;padding:10px 12px;text-align:left;font-weight:600}
  td{border-bottom:1px solid #e2e8f0;vertical-align:middle}
  tr:last-child td{border-bottom:none}
  .disclaimer{background:#fff8e1;border:1px solid #ffe082;border-radius:8px;padding:14px 18px;font-size:0.82em;color:#555;line-height:1.6}
  .footer{text-align:center;color:#999;font-size:0.78em;padding:16px 0 32px}
  @media(max-width:600px){.score-box{flex-direction:column}.kpi-grid{grid-template-columns:1fr 1fr}}
</style>
</head>
<body>
<div class="header">
  <h1>&#x1F6E1; Valutazione Sicurezza Microsoft 365</h1>
  <div class="sub">Tenant: <strong>$TenantName</strong> &nbsp;|&nbsp; Data: $dateStr &nbsp;|&nbsp; Periodo analisi: $LookbackDays giorni &nbsp;|&nbsp; Versione script: v$($Script:AssessmentVersion)</div>
</div>
<div class="container">

  <!-- SCORE -->
  <div class="card">
    <div class="card-header">&#x2B50; Security Score Complessivo</div>
    <div class="card-body">
      <div class="score-box">
        <div class="score-circle">
          <div class="score-num">$pct%</div>
          <div class="score-label">$rating</div>
        </div>
        <div class="score-details">
          <h2>$rating</h2>
          <div class="sub">Punteggio: $score / $maxScore punti totali</div>
          <div class="gauge-wrap"><div class="gauge-fill"></div></div>
          <div class="gauge-lbl">$pct% del massimo raggiunto</div>
        </div>
      </div>
    </div>
  </div>

  <!-- KPI -->
  <div class="card">
    <div class="card-header">&#x1F4CA; KPI Chiave</div>
    <div class="card-body">
      <div class="kpi-grid">
        <div class="kpi"><div class="val">$mfaCov%</div><div class="lbl">Copertura MFA</div></div>
        <div class="kpi"><div class="val">$enabledCa</div><div class="lbl">Policy CA Attive</div></div>
        <div class="kpi"><div class="val">$adminCount</div><div class="lbl">Global Administrators</div></div>
        <div class="kpi"><div class="val">$highRiskOAuth</div><div class="lbl">OAuth Alto Rischio</div></div>
        <div class="kpi"><div class="val">$legacyCount</div><div class="lbl">Eventi Legacy Auth</div></div>
        <div class="kpi"><div class="val">$riskyFwdCount</div><div class="lbl">Forward Sospetti</div></div>
        <div class="kpi"><div class="val">$smtpAuthStatus</div><div class="lbl">SMTP Auth Globale</div></div>
        <div class="kpi"><div class="val">$adminSenzaMfa</div><div class="lbl">Admin senza MFA</div></div>
      </div>
    </div>
  </div>

  <!-- RISCHI -->
  <div class="card">
    <div class="card-header">&#x26A0; Rischi Principali</div>
    <div class="card-body" style="padding:0">
      <table>
        <thead><tr><th>Area</th><th>Stato</th><th>Livello</th><th>Note</th></tr></thead>
        <tbody>$riskRows</tbody>
      </table>
    </div>
  </div>

  <!-- AZIONI -->
  <div class="card">
    <div class="card-header">&#x1F3AF; Azioni Prioritarie Raccomandate</div>
    <div class="card-body" style="padding:0">
      <table>
        <thead><tr><th>Priorita</th><th>Categoria</th><th>Azione</th><th>Scadenza</th></tr></thead>
        <tbody>$actionRows</tbody>
      </table>
    </div>
  </div>

  <!-- DISCLAIMER -->
  <div class="disclaimer">
    <strong>&#x26A0; Dichiarazione di limitazione:</strong> Questo report e stato generato automaticamente da uno script read-only. I dati riflettono lo stato del tenant al momento dell'assessment (<em>$dateStr</em>). Nessuna modifica e stata apportata al tenant. Le raccomandazioni devono essere validate da un esperto di sicurezza prima dell'implementazione. Il report non contiene dati sensibili personali estesi.
  </div>

  <div class="footer">Report M365 Security Assessment v$($Script:AssessmentVersion) &bull; $TenantName &bull; $dateFile</div>
</div>
</body>
</html>
"@

    $html | Set-Content -Path $HtmlPath -Encoding UTF8
    Write-Progress -Activity "Generazione HTML Report" -Completed
    Write-Log "HTML Report esecutivo salvato: $HtmlPath" "SUCCESS"
    return $HtmlPath
}

#endregion

#region ─── MAIN ───────────────────────────────────────────────────────────────

function Main {
    $Script:OutputDir = $OutputPath
    if (-not (Test-Path $Script:OutputDir)) {
        New-Item -ItemType Directory -Path $Script:OutputDir -Force | Out-Null
    }
    $csvDir = Join-Path $Script:OutputDir "CSV"
    if (-not (Test-Path $csvDir)) { New-Item -ItemType Directory -Path $csvDir -Force | Out-Null }

    $logPath = Join-Path $Script:OutputDir "assessment.log"
    Initialize-Logging -LogPath $logPath

    Write-Log "Output directory : $Script:OutputDir" "INFO"
    Write-Log "CSV directory    : $csvDir" "INFO"
    Write-Log "SkipGraph        : $SkipGraph" "INFO"
    Write-Log "SkipExchange     : $SkipExchange" "INFO"

    try {
        # ─ Fase 1: Moduli ──────────────────────────────────────────────────────
        Write-Progress -Activity "M365 Security Assessment v$Script:AssessmentVersion" `
            -Status "Fase 1/7: Verifica moduli..." -PercentComplete 5
        Test-AndImportModules

        # ─ Fase 2: Connessioni ─────────────────────────────────────────────────
        Write-Progress -Activity "M365 Security Assessment v$Script:AssessmentVersion" `
            -Status "Fase 2/7: Connessione servizi..." -PercentComplete 12
        if (-not $SkipGraph)    { Connect-ToMicrosoftGraph }
        if (-not $SkipExchange) { Connect-ToExchangeOnline }

        # ─ Fase 3: Raccolta Graph ──────────────────────────────────────────────
        Write-Progress -Activity "M365 Security Assessment v$Script:AssessmentVersion" `
            -Status "Fase 3/7: Raccolta dati Microsoft Graph..." -PercentComplete 20
        Write-Section "RACCOLTA DATI GRAPH"

        $mfaData            = if (-not $SkipGraph)    { Get-MFARegistrationData }              else { $null }
        $caData             = if (-not $SkipGraph)    { Get-ConditionalAccessData }            else { $null }
        $signInData         = if (-not $SkipGraph)    { Get-SignInLogsData -Days $LookbackDays }else { $null }
        $legacyData         = if ($signInData)         { Get-LegacyAuthData -SignInLogs $signInData } else { $null }
        $rolesData          = if (-not $SkipGraph)    { Get-DirectoryRolesData }               else { $null }
        $oauthData          = if (-not $SkipGraph)    { Get-OAuthGrantsData }                  else { $null }
        $riskyUsersData     = if (-not $SkipGraph)    { Get-RiskyUsersData }                   else { $null }
        $riskDetectionsData = if (-not $SkipGraph)    { Get-RiskDetectionsData }               else { $null }

        # ─ Fase 4: Raccolta Exchange ───────────────────────────────────────────
        Write-Progress -Activity "M365 Security Assessment v$Script:AssessmentVersion" `
            -Status "Fase 4/7: Raccolta dati Exchange Online..." -PercentComplete 45
        Write-Section "RACCOLTA DATI EXCHANGE"

        $transportData = if (-not $SkipExchange) { Get-TransportConfigData }  else { $null }
        $casData       = if (-not $SkipExchange) { Get-CASMailboxData }       else { $null }
        $mailboxData   = if (-not $SkipExchange) { Get-MailboxData }          else { $null }
        $rulesData     = if (-not $SkipExchange -and $mailboxData) {
            Get-MailboxRulesData -Mailboxes $mailboxData
        } else { $null }

        # ─ Fase 5: Score ───────────────────────────────────────────────────────
        Write-Progress -Activity "M365 Security Assessment v$Script:AssessmentVersion" `
            -Status "Fase 5/7: Calcolo Security Score..." -PercentComplete 62
        Write-Section "CALCOLO SECURITY SCORE"
        $scoreResult = Get-SecurityScore `
            -TransportConfig $transportData `
            -MfaData         $mfaData `
            -CaPolicies      $caData `
            -DirectoryRoles  $rolesData `
            -CasMailboxes    $casData `
            -OAuthGrants     $oauthData

        Write-Log "Security Score: $($scoreResult.Score)/$($scoreResult.MaxScore) ($($scoreResult.Percentage)%) - $($scoreResult.Rating)" "SUCCESS"

        # ─ Fase 6: Export CSV + Excel + Report ────────────────────────────────
        Write-Progress -Activity "M365 Security Assessment v$Script:AssessmentVersion" `
            -Status "Fase 6/7: Export CSV..." -PercentComplete 68
        $csvFiles = Export-DataToCSV -CsvDir $csvDir `
            -MfaData            $mfaData `
            -CaData             $caData `
            -SignInData          $signInData `
            -LegacyData         $legacyData `
            -TransportData      $transportData `
            -CasData            $casData `
            -MailboxData        $mailboxData `
            -RolesData          $rolesData `
            -OAuthData          $oauthData `
            -RiskyUsersData     $riskyUsersData `
            -RiskDetectionsData $riskDetectionsData `
            -RulesData          $rulesData

        Write-Progress -Activity "M365 Security Assessment v$Script:AssessmentVersion" `
            -Status "Fase 6/7: Generazione Excel Workbook..." -PercentComplete 76
        $dateStamp = Get-Date -Format "yyyyMMdd"
        $excelPath = Join-Path $Script:OutputDir "M365_SecurityAssessment_${TenantName}_${dateStamp}.xlsx"
        $excelFile = New-ExcelWorkbook -ExcelPath $excelPath `
            -ScoreResult        $scoreResult `
            -MfaData            $mfaData `
            -CaData             $caData `
            -SignInData          $signInData `
            -LegacyData         $legacyData `
            -TransportData      $transportData `
            -CasData            $casData `
            -MailboxData        $mailboxData `
            -RolesData          $rolesData `
            -OAuthData          $oauthData `
            -RiskyUsersData     $riskyUsersData `
            -RiskDetectionsData $riskDetectionsData `
            -RulesData          $rulesData

        Write-Progress -Activity "M365 Security Assessment v$Script:AssessmentVersion" `
            -Status "Fase 6/7: Generazione Report Markdown..." -PercentComplete 84
        $reportPath = Join-Path $Script:OutputDir "RapportoEsecutivo_${TenantName}_${dateStamp}.md"
        $reportFile = New-ExecutiveReport -ReportPath $reportPath `
            -ScoreResult  $scoreResult `
            -MfaData      $mfaData `
            -CaData       $caData `
            -TransportData $transportData `
            -RolesData    $rolesData `
            -LegacyData   $legacyData `
            -OAuthData    $oauthData `
            -RulesData    $rulesData

        Write-Progress -Activity "M365 Security Assessment v$Script:AssessmentVersion" `
            -Status "Fase 6/7: Generazione HTML Report esecutivo..." -PercentComplete 90
        $htmlPath = Join-Path $Script:OutputDir "M365_SecurityAssessment_${TenantName}_${dateStamp}.html"
        $htmlFile = New-HtmlReport -HtmlPath $htmlPath `
            -ScoreResult  $scoreResult `
            -MfaData      $mfaData `
            -CaData       $caData `
            -TransportData $transportData `
            -RolesData    $rolesData `
            -LegacyData   $legacyData `
            -OAuthData    $oauthData `
            -RulesData    $rulesData `
            -CasData      $casData

        # ─ Fase 7: Riepilogo ───────────────────────────────────────────────────
        Write-Progress -Activity "M365 Security Assessment v$Script:AssessmentVersion" `
            -Status "Fase 7/7: Riepilogo finale..." -PercentComplete 97
        Write-Section "ASSESSMENT COMPLETATO"
        $duration = (Get-Date) - $Script:StartTime
        Write-Log "Durata totale: $([math]::Round($duration.TotalMinutes, 1)) minuti" "INFO"
        Write-Log "" "INFO"
        Write-Log "SECURITY SCORE: $($scoreResult.Score)/$($scoreResult.MaxScore) ($($scoreResult.Percentage)%) - $($scoreResult.Rating)" "SUCCESS"
        Write-Log "" "INFO"
        Write-Log "FILE GENERATI:" "INFO"
        Write-Log "  Excel Workbook  : $excelFile"  "SUCCESS"
        Write-Log "  HTML Report     : $htmlFile"   "SUCCESS"
        Write-Log "  Report Esec. MD : $reportFile" "SUCCESS"
        Write-Log "  CSV Directory   : $csvDir"     "SUCCESS"
        Write-Log "  Log File        : $logPath"    "SUCCESS"
        Write-Log "" "INFO"
        Write-Log "CSV Files ($($csvFiles.Count)):" "INFO"
        foreach ($csv in $csvFiles) { Write-Log "  - $(Split-Path $csv -Leaf)" "INFO" }

        Write-Progress -Activity "M365 Security Assessment v$Script:AssessmentVersion" -Completed

        [PSCustomObject]@{
            TenantName      = $TenantName
            SecurityScore   = "$($scoreResult.Score)/$($scoreResult.MaxScore)"
            Percentage      = "$($scoreResult.Percentage)%"
            Rating          = $scoreResult.Rating
            ExcelReport     = $excelFile
            HtmlReport      = $htmlFile
            ExecutiveReport = $reportFile
            CSVDirectory    = $csvDir
            LogFile         = $logPath
            Duration        = "$([math]::Round($duration.TotalMinutes, 1)) min"
            AssessmentDate  = (Get-Date -Format "yyyy-MM-dd HH:mm")
        }

    } catch {
        Write-Log "ERRORE CRITICO: $($_.Exception.Message)" "ERROR"
        Write-Log "Stack trace: $($_.ScriptStackTrace)" "ERROR"
        throw
    } finally {
        Write-Log "Disconnessione servizi in corso..." "INFO"
        try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
        Write-Log "Disconnessione completata." "INFO"
    }
}

#region ─── ENTRY POINT ───────────────────────────────────────────────────────

# -SyntaxOnly: esegue solo la verifica del parser e termina senza connettersi al tenant.
# Utile per CI/CD, pre-deploy checks, o debug offline.
if ($SyntaxOnly) {
    Write-Host ""
    Write-Host "  -SyntaxOnly attivo: verifica sintattica del file in corso..." -ForegroundColor Cyan
    Write-Host ""
    $syntaxResult = Test-ScriptSyntax -Path $PSCommandPath
    if ($syntaxResult.IsValid) {
        Write-Host ""
        Write-Host "  Versione: $Script:AssessmentVersion  |  File: $PSCommandPath" -ForegroundColor DarkGray
        exit 0
    } else {
        exit 1
    }
}

Main

#endregion