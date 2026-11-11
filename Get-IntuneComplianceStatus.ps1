<#
.SYNOPSIS
    Shows the compliance state of all Intune managed devices – including the reasons for
    noncompliance – and creates an HTML tile report.

.DESCRIPTION
    Reads all managed devices via Microsoft Graph (beta) and, for every device that is not
    compliant, the per-policy and per-setting compliance states:
      - Devices                     /deviceManagement/managedDevices
      - Policy + setting states     /deviceManagement/managedDevices/{id}/deviceCompliancePolicyStates

    If Graph reports a policy as noncompliant but returns no setting details for it, the script
    tries the reporting API the Intune admin center uses for the same view:
      - POST /deviceManagement/reports/getDevicePolicySettingsComplianceReport
    If that does not return usable data either, the tile shows the policy with a hint to check
    the device in the admin center.

    The HTML report contains one tile per device (device name, user, OS, status label):
      green  = compliant
      red    = not compliant (noncompliant, error, conflict)
      orange = in grace period
      grey   = unknown / not evaluated / reported by ConfigMgr
    Clicking a tile opens the details: device data, the failing compliance policies and the
    failing settings with current value and error code (if reported).
    The HTML file is self-contained (no external resources) and can be sent by mail.

    Per-device details are resolved via JSON batching ($batch, 20 requests per call,
    retry on HTTP 429/503/504). Errors while loading details do not abort the run; they are
    reported as warnings, summarized at the end and shown as a banner in the report.

    Required permission (delegated or application):
      DeviceManagementManagedDevices.Read.All (or ReadWrite.All)

    Sign-in:
      Without an existing connection the script signs in interactively (delegated).
      For app-only (e.g. automation) connect yourself beforehand, e.g.:
        Connect-MgGraph -TenantId <tenant> -ClientId <appId> -CertificateThumbprint <thumbprint>
      The script then verifies the application permission with a test request.

.PARAMETER TenantId
    Optional: tenant ID or any verified domain of the tenant. Forces the connection to exactly this tenant
    (important for multi-tenant admins / guest accounts).

.PARAMETER DeviceName
    Optional: only devices whose name matches the pattern, e.g. "WS-*". Wildcards * and ? are
    supported; text without wildcards is treated as "name contains", e.g. "WS" -> "*WS*".
    If omitted in an interactive session, the script asks for it (Enter = all devices).

.PARAMETER OperatingSystem
    Optional: only devices of the given platform(s): Windows, iOS (incl. iPadOS), Android, macOS, Linux.
    If omitted in an interactive session, the script shows a selection menu (Enter = all platforms).

.PARAMETER NoPrompt
    Never ask for filters; without -DeviceName / -OperatingSystem all devices are reported.
    Prompts are skipped automatically for an existing app-only connection and for non-interactive
    hosts (e.g. scheduled tasks, Azure Automation, pwsh -NonInteractive). The prompts come BEFORE
    the sign-in, so a run can still be aborted without authenticating.

.PARAMETER OnlyNonCompliant
    Only list devices that are not compliant (everything except "compliant").

.PARAMETER IncludeCompliantDetails
    Also load the policy states for compliant devices (shows the evaluated policies in the tile
    details). Costs one additional Graph request per device – not recommended for large tenants.

.PARAMETER OutputPath
    Path of the HTML report. The target folder is created if missing.
    Default: <TEMP>\Intune-ComplianceStatus_<timestamp>.html

.PARAMETER CsvPath
    Optional: additionally export one row per device (incl. reasons) as CSV (UTF-8 with BOM).

.PARAMETER Delimiter
    CSV delimiter. Default ';' (opens directly in Excel with German/European regional settings).

.PARAMETER Open
    Open the HTML report in the default browser when finished.

.PARAMETER PassThru
    Return the result objects to the pipeline.

.EXAMPLE
    .\Get-IntuneComplianceStatus.ps1 -TenantId contoso.onmicrosoft.com -Open
    Asks for name filter and platforms, then opens the report.
.EXAMPLE
    .\Get-IntuneComplianceStatus.ps1 -NoPrompt -CsvPath C:\Reports\all-devices.csv
.EXAMPLE
    .\Get-IntuneComplianceStatus.ps1 -OperatingSystem Windows -OnlyNonCompliant -CsvPath C:\Reports\noncompliant.csv
.EXAMPLE
    .\Get-IntuneComplianceStatus.ps1 -DeviceName "WS-*" -IncludeCompliantDetails -Open
.EXAMPLE
    $r = .\Get-IntuneComplianceStatus.ps1 -PassThru; $r | Where-Object Status -eq 'noncompliant' | Select-Object DeviceName, Reasons

.NOTES
    Author:   Timo Schlappinger
    Version:  1.1.2
    Project:  https://github.com/tim2611/Intune-Assignment-Report
    Requires: Module Microsoft.Graph.Authentication, PowerShell 5.1 or 7.x
#>
[CmdletBinding()]
[OutputType([pscustomobject])]
param(
    [string]$TenantId,
    [string]$DeviceName,
    [ValidateSet('Windows', 'iOS', 'Android', 'macOS', 'Linux')][string[]]$OperatingSystem,
    [switch]$OnlyNonCompliant,
    [switch]$IncludeCompliantDetails,
    [string]$OutputPath = (Join-Path ([IO.Path]::GetTempPath()) "Intune-ComplianceStatus_$(Get-Date -Format 'yyyyMMdd_HHmm').html"),
    [string]$CsvPath,
    [ValidateSet(';', ',', "`t")][string]$Delimiter = ';',
    [switch]$Open,
    [switch]$NoPrompt,
    [switch]$PassThru
)

$ErrorActionPreference = 'Stop'
$graphBeta  = 'https://graph.microsoft.com/beta'
$loadErrors = [System.Collections.Generic.List[string]]::new()
$osOptions  = @('Windows', 'iOS', 'Android', 'macOS', 'Linux')   # keep in sync with ValidateSet of -OperatingSystem and $osPattern

# Effective filters (separate variables: the parameters carry validation attributes)
$nameFilter = $DeviceName
$osFilter   = @($OperatingSystem | Where-Object { $_ })

#region Connection & permissions
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    throw "Module missing: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
}
Import-Module Microsoft.Graph.Authentication
#endregion

#region Interactive filter selection (only for parameters that were not passed)
function Read-OsSelection {
    param([string[]]$Options)
    Write-Host ""
    Write-Host "Which operating systems should be included?" -ForegroundColor Cyan
    for ($i = 0; $i -lt $Options.Count; $i++) {
        $suffix = if ($Options[$i] -eq 'iOS') { ' (incl. iPadOS)' } else { '' }
        Write-Host "  [$($i + 1)] $($Options[$i])$suffix"
    }
    while ($true) {
        $answer = Read-Host "Selection, e.g. 1,3 or Windows,iOS (Enter = all)"
        if ([string]::IsNullOrWhiteSpace($answer)) { return @() }
        $tokens = @($answer -split '[,;\s]+' | Where-Object { $_ })
        $picked = [System.Collections.Generic.List[string]]::new()
        $valid  = $true
        foreach ($t in $tokens) {
            $n = 0
            $name = if ([int]::TryParse($t, [ref]$n) -and $n -ge 1 -and $n -le $Options.Count) { $Options[$n - 1] }
                    else { $Options | Where-Object { $_ -eq $t } | Select-Object -First 1 }   # names, case-insensitive
            if ($name) { if (-not $picked.Contains($name)) { $picked.Add($name) } } else { $valid = $false }
        }
        if ($valid -and $picked.Count -gt 0) { return @($picked) }
        Write-Warning "Invalid selection '$answer' – use numbers 1-$($Options.Count) or names, separated by commas."
    }
}

# The prompts deliberately run BEFORE Connect-MgGraph: the filters do not depend on the tenant, and
# the user can decide (or abort) before the sign-in. Consequently only an EXISTING app-only connection
# can suppress them here; a new connection made by this script is always delegated, i.e. interactive.
# The host check covers the usual automation cases; '^[-/]noni' deliberately also matches the
# abbreviations PowerShell accepts for -NonInteractive (e.g. "pwsh -noni"). Hosts that still cannot
# read input are caught below: Read-Host then throws and the run continues without filter.
$existingCtx = Get-MgContext
$canPrompt   = -not $NoPrompt -and
               [Environment]::UserInteractive -and
               -not ([Environment]::GetCommandLineArgs() | Where-Object { $_ -match '^[-/]noni' }) -and
               -not ($existingCtx -and "$($existingCtx.AuthType)" -eq 'AppOnly')

if ($canPrompt) {
    try {
        if (-not $PSBoundParameters.ContainsKey('DeviceName')) {
            Write-Host ""
            $nameFilter = "$(Read-Host "Device name filter (wildcards * and ?, text without wildcard = 'contains', Enter = all)")".Trim()
        }
        if (-not $PSBoundParameters.ContainsKey('OperatingSystem')) {
            $osFilter = @(Read-OsSelection -Options $osOptions)
        }
    }
    catch {
        Write-Warning "This host cannot read input ($($_.Exception.Message)) – continuing without interactive filters."
        $nameFilter = $DeviceName
        $osFilter   = @($OperatingSystem | Where-Object { $_ })
    }
}

# Text without wildcards means "contains"
if ($nameFilter -and $nameFilter -notmatch '[\*\?]') { $nameFilter = "*$nameFilter*" }

$filterText = @(
    if ($nameFilter)       { "name like '$nameFilter'" }
    if ($osFilter.Count)   { "OS: $($osFilter -join ', ')" }
    if ($OnlyNonCompliant) { 'only not compliant' }
)
Write-Host "Filter: $(if ($filterText.Count) { $filterText -join '; ' } else { 'none (all devices)' })" -ForegroundColor Cyan
#endregion

#region Connection & permissions (continued)
# Required permission -> accepted (equal or higher) permissions
$scopeMap = [ordered]@{
    'DeviceManagementManagedDevices.Read.All' = @('DeviceManagementManagedDevices.Read.All', 'DeviceManagementManagedDevices.ReadWrite.All')
}

function Get-MissingScopes {
    $ctx = Get-MgContext
    if (-not $ctx) { return @($scopeMap.Keys) }
    # Normalization: if scopes are returned as full URLs, use the short name only
    $have = @($ctx.Scopes | ForEach-Object { ($_ -split '/')[-1] })
    $missing = foreach ($required in $scopeMap.Keys) {
        if (-not ($scopeMap[$required] | Where-Object { $_ -in $have })) { $required }
    }
    return @($missing)
}

function Resolve-TenantGuid {
    # -TenantId may be a GUID or a domain; Get-MgContext always returns the GUID.
    # Domains are resolved anonymously via the OpenID configuration of the tenant.
    param([string]$Tenant)
    $g = [guid]::Empty
    if ([guid]::TryParse($Tenant, [ref]$g)) { return $g.ToString() }
    try {
        $cfg = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$([uri]::EscapeDataString($Tenant))/v2.0/.well-known/openid-configuration" -ErrorAction Stop
        if ("$($cfg.issuer)" -match '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})') { return $Matches[1].ToLower() }
    }
    catch { }
    return $null
}

$tenantGuid = $null
if ($TenantId) {
    $tenantGuid = Resolve-TenantGuid $TenantId
    if (-not $tenantGuid) { Write-Warning "Tenant '$TenantId' could not be resolved to a tenant ID – the existing connection is not checked against it." }
}

$ctx         = Get-MgContext
$isAppOnly   = $ctx -and ("$($ctx.AuthType)" -eq 'AppOnly')
$wrongTenant = $ctx -and $tenantGuid -and ("$($ctx.TenantId)" -ne $tenantGuid)

if ($isAppOnly) {
    if ($wrongTenant) {
        throw "Existing app connection points to tenant $($ctx.TenantId), not $TenantId. Run Disconnect-MgGraph and reconnect with the matching app registration."
    }
}
elseif (-not $ctx -or $wrongTenant -or (Get-MissingScopes).Count -gt 0) {
    $connectParams = @{ Scopes = @($scopeMap.Keys); NoWelcome = $true }
    if ($TenantId) { $connectParams.TenantId = $TenantId }
    Connect-MgGraph @connectParams
}

$ctx = Get-MgContext
if ("$($ctx.AuthType)" -eq 'AppOnly') {
    # Application permissions are not reliably visible in the context -> functional test via probe request
    try {
        Invoke-MgGraphRequest -Method GET -Uri "$graphBeta/deviceManagement/managedDevices?`$top=1&`$select=id" | Out-Null
    }
    catch {
        throw "App registration is (probably) missing the application permission DeviceManagementManagedDevices.Read.All (incl. admin consent): $($_.Exception.Message)"
    }
}
else {
    $missing = Get-MissingScopes
    if ($missing.Count -gt 0) {
        throw "Missing delegated permissions: $($missing -join ', ')"
    }
}
$identity = if ($ctx.Account) { $ctx.Account } else { $ctx.AppName }
Write-Host "Connected to tenant $($ctx.TenantId) as $identity ($($ctx.AuthType))" -ForegroundColor Cyan
#endregion

#region Helper functions
function Get-GraphAll {
    param([Parameter(Mandatory)][string]$Uri)
    $items = [System.Collections.Generic.List[object]]::new()
    do {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $Uri
        foreach ($v in $resp.value) { $items.Add($v) }
        $Uri = $resp.'@odata.nextLink'
    } while ($Uri)
    return $items
}

function Invoke-GraphBatch {
    <#
      Runs GET requests via JSON batching (max. 20 per call).
      Single 429/503/504 responses inside a batch are retried up to 3 times honoring Retry-After,
      as are 429/5xx/transport errors of the batch POST itself. If the POST of one chunk still fails,
      its requests get status 0 and the remaining chunks continue (results already loaded are kept).
      4xx errors of the POST itself (e.g. 401/403) abort immediately – they would hit every chunk.
      Returns: hashtable relativeUrl -> response object (status, body).
    #>
    param(
        [Parameter(Mandatory)][string[]]$RelativeUrls,
        [ValidateSet('v1.0', 'beta')][string]$Version = 'v1.0'
    )
    $maxAttempts = 4
    $retryable   = @(429, 503, 504)
    $batchUri    = "https://graph.microsoft.com/$Version/`$batch"
    $results     = @{}

    for ($i = 0; $i -lt $RelativeUrls.Count; $i += 20) {
        $pending  = @($RelativeUrls[$i..([Math]::Min($i + 19, $RelativeUrls.Count - 1))])
        $lastResp = @{}

        for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
            $requests = for ($j = 0; $j -lt $pending.Count; $j++) {
                @{ id = "$j"; method = 'GET'; url = $pending[$j] }
            }
            $body = @{ requests = @($requests) } | ConvertTo-Json -Depth 5 -Compress

            # Protect the batch POST itself: 5xx/transport errors -> retry, 4xx (e.g. 401/403) -> abort immediately
            try {
                $resp = Invoke-MgGraphRequest -Method POST -Uri $batchUri -Body $body -ContentType 'application/json'
            }
            catch {
                $code = $null
                if ($_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode }
                if ($code -and $code -lt 500 -and $code -ne 429) { throw }
                if ($attempt -eq $maxAttempts) {
                    Write-Warning "Batch request failed after $maxAttempts attempts: $($_.Exception.Message)"
                    break
                }
                Start-Sleep -Seconds (5 * $attempt)
                continue
            }

            $retry = [System.Collections.Generic.List[string]]::new()
            $wait  = 0
            foreach ($r in $resp.responses) {
                $url = $pending[[int]$r.id]
                if ([int]$r.status -in $retryable) {
                    $retry.Add($url)
                    $lastResp[$url] = $r
                    $ra  = 0
                    $raw = if ($r.headers) { @($r.headers['Retry-After'])[0] }
                    if ($raw -and [int]::TryParse("$raw".Trim(), [ref]$ra) -and $ra -gt $wait) { $wait = $ra }
                }
                else {
                    $results[$url] = $r
                }
            }

            $pending = @($retry)
            if ($pending.Count -eq 0) { break }
            if ($attempt -lt $maxAttempts) { Start-Sleep -Seconds ([Math]::Max($wait, 5)) }
        }

        # Requests still open after all attempts: keep their last response (status)
        foreach ($url in $pending) {
            $results[$url] = if ($lastResp.ContainsKey($url)) { $lastResp[$url] } else { @{ status = 0; body = $null } }
        }
    }
    return $results
}

function ConvertTo-IsoDate {
    # Graph dates may arrive as string or DateTime; "0001-01-01" means "never"
    param($Value)
    if ($null -eq $Value -or "$Value" -eq '') { return $null }
    $dt = [datetime]::MinValue
    if ($Value -is [datetime]) { $dt = $Value }
    elseif (-not [datetime]::TryParse("$Value", [Globalization.CultureInfo]::InvariantCulture,
                                      [Globalization.DateTimeStyles]::RoundtripKind, [ref]$dt)) { return "$Value" }
    if ($dt.Year -le 1) { return $null }
    if ($dt.Kind -eq [DateTimeKind]::Unspecified) { $dt = [datetime]::SpecifyKind($dt, [DateTimeKind]::Utc) }
    return $dt.ToUniversalTime().ToString('o')
}

function Get-StatusClass {
    # Maps managedDevice.complianceState to the four tile colors
    param([string]$State)
    switch ($State) {
        'compliant'                                       { return 'compliant' }
        { $_ -in 'noncompliant', 'conflict', 'error' }    { return 'noncompliant' }
        'inGracePeriod'                                   { return 'grace' }
        default                                           { return 'unknown' }
    }
}

# German display names for common compliance settings (key = last segment of the setting id,
# lookup is case-insensitive). Unknown settings fall back to the name Graph returns.
$settingLabels = @{
    # Default Device Compliance Policy (built-in)
    'RequireDeviceCompliancePolicyAssigned'      = 'Keine Compliance-Richtlinie zugewiesen'
    'RequireRemainContact'                       = 'Gerät hat sich nicht innerhalb der Gültigkeitsdauer gemeldet'
    'RequireUserExistence'                       = 'Registrierter Benutzer existiert nicht (mehr)'
    # Windows – Device Health
    'BitLockerEnabled'                           = 'BitLocker erforderlich'
    'SecureBootEnabled'                          = 'Secure Boot erforderlich'
    'CodeIntegrityEnabled'                       = 'Codeintegrität erforderlich'
    'EarlyLaunchAntiMalwareDriverEnabled'        = 'ELAM-Treiber erforderlich'
    # Device properties
    'OsMinimumVersion'                           = 'Minimale Betriebssystemversion'
    'OsMaximumVersion'                           = 'Maximale Betriebssystemversion'
    'MobileOsMinimumVersion'                     = 'Minimale Betriebssystemversion (mobil)'
    'MobileOsMaximumVersion'                     = 'Maximale Betriebssystemversion (mobil)'
    'MinAndroidSecurityPatchLevel'               = 'Minimale Android-Sicherheitspatchebene'
    'OsMinimumBuildVersion'                      = 'Minimale Build-Version'
    'OsMaximumBuildVersion'                      = 'Maximale Build-Version'
    # System security
    'StorageRequireEncryption'                   = 'Verschlüsselung des Datenspeichers'
    'ActiveFirewallRequired'                     = 'Firewall erforderlich'
    'FirewallEnabled'                            = 'Firewall erforderlich'
    'TpmRequired'                                = 'TPM erforderlich'
    'AntivirusRequired'                          = 'Antivirus erforderlich'
    'AntiSpywareRequired'                        = 'Antispyware erforderlich'
    'DefenderEnabled'                            = 'Microsoft Defender Antimalware'
    'DefenderVersion'                            = 'Minimale Version Defender Antimalware'
    'SignatureOutOfDate'                         = 'Defender-Signaturen aktuell'
    'RtpEnabled'                                 = 'Echtzeitschutz'
    'ConfigurationManagerComplianceRequired'     = 'Compliance aus Configuration Manager'
    'PasswordRequired'                           = 'Kennwort erforderlich'
    'PasswordMinimumLength'                      = 'Minimale Kennwortlänge'
    'PasswordRequiredType'                       = 'Kennworttyp'
    'PasswordMinutesOfInactivityBeforeLock'      = 'Sperre nach Inaktivität'
    'PasswordExpirationDays'                     = 'Kennwortablauf'
    'PasswordPreviousPasswordBlockCount'         = 'Kennwortverlauf'
    'PasswordBlockSimple'                        = 'Einfache Kennwörter blockieren'
    'PasscodeRequired'                           = 'Kennung erforderlich'
    'SecurityBlockJailbrokenDevices'             = 'Gerät mit Jailbreak/Root blockiert'
    'SecurityPreventInstallAppsFromUnknownSources' = 'Apps aus unbekannten Quellen blockiert'
    'SecurityRequireSafetyNetAttestationBasicIntegrity' = 'Play Integrity: Basisintegrität'
    'SecurityRequireSafetyNetAttestationCertifiedDevice' = 'Play Integrity: zertifiziertes Gerät'
    'SystemIntegrityProtectionEnabled'           = 'System Integrity Protection (macOS)'
    'GatekeeperAllowedAppSource'                 = 'Gatekeeper'
    # Defender for Endpoint / Mobile Threat Defense
    'DeviceThreatProtectionEnabled'              = 'Bedrohungsschutz (Mobile Threat Defense)'
    'DeviceThreatProtectionRequiredSecurityLevel' = 'Maximal zulässige Gerätebedrohungsstufe'
    'DeviceCompliancePolicyScript'               = 'Benutzerdefinierte Compliance (Skript)'
}

function Get-SettingLabel {
    param([string]$Setting, [string]$SettingName)
    foreach ($candidate in $Setting, $SettingName) {
        $key = ("$candidate" -split '\.')[-1]
        if ($key -and $settingLabels.ContainsKey($key)) { return $settingLabels[$key] }
    }
    if ($SettingName) { return $SettingName }
    if ($Setting)     { return $Setting }
    return '<unknown setting>'
}

$failStates = @('nonCompliant', 'error', 'conflict')

function Get-SettingFallback {
    <#
      Fallback for policies that are reported as noncompliant without setting details:
      the reporting API behind the admin center view "device > device compliance > policy".
      The column layout of this report is not documented; columns are resolved by name.
      Returns nothing if the report cannot be read or interpreted.
    #>
    param([string]$DeviceId, [string]$PolicyId)

    # Both values go into an OData filter -> only accept GUIDs
    $g = [guid]::Empty
    if (-not ([guid]::TryParse($DeviceId, [ref]$g) -and [guid]::TryParse($PolicyId, [ref]$g))) { return }

    $body   = @{ filter = "(DeviceId eq '$DeviceId') and (PolicyId eq '$PolicyId')" } | ConvertTo-Json -Compress
    $report = $null
    for ($attempt = 1; $attempt -le 3 -and -not $report; $attempt++) {
        $tmp = [IO.Path]::GetTempFileName()
        try {
            Invoke-MgGraphRequest -Method POST -Uri "$graphBeta/deviceManagement/reports/getDevicePolicySettingsComplianceReport" `
                -Body $body -ContentType 'application/json' -OutputFilePath $tmp | Out-Null
            $report = [IO.File]::ReadAllText($tmp, [Text.Encoding]::UTF8) | ConvertFrom-Json
        }
        catch {
            $code = $null
            if ($_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode }
            # Retry only throttling / server errors; anything else (e.g. 400 for an unsupported policy type) is final
            if ($attempt -eq 3 -or ($code -and $code -lt 500 -and $code -ne 429)) {
                Write-Verbose "Fallback report failed for device $DeviceId / policy $PolicyId : $($_.Exception.Message)"
                return
            }
            Start-Sleep -Seconds (5 * $attempt)
        }
        finally {
            Remove-Item -Path $tmp -Force -ErrorAction SilentlyContinue
        }
    }
    if (-not $report) { return }

    $cols      = @($report.Schema | ForEach-Object { "$($_.Column)" })
    $nameCol   = @('SettingNm', 'SettingName', 'Setting')         | Where-Object { $_ -in $cols } | Select-Object -First 1
    $rawCol    = @('SettingStatus', 'ComplianceState')            | Where-Object { $_ -in $cols } | Select-Object -First 1
    $locCol    = @('SettingStatus_loc', 'ComplianceState_loc')    | Where-Object { $_ -in $cols } | Select-Object -First 1
    $valueCol  = @('CurrentValue', 'Value')                       | Where-Object { $_ -in $cols } | Select-Object -First 1
    $errorCol  = @('ErrorCode')                                   | Where-Object { $_ -in $cols } | Select-Object -First 1
    if (-not $nameCol -or -not ($rawCol -or $locCol)) { return }

    foreach ($v in @($report.Values)) {
        $row = @{}
        for ($i = 0; $i -lt $cols.Count; $i++) { $row[$cols[$i]] = @($v)[$i] }

        # Positive list: only clearly failing states count. The machine-readable column is preferred;
        # the localized column is only used if the raw value is not a known state name (e.g. numeric).
        $raw = if ($rawCol) { "$($row[$rawCol])" } else { '' }
        $loc = if ($locCol) { "$($row[$locCol])" } else { '' }
        if ($raw -match '^[A-Za-z]+$') {
            if ($raw -notin $failStates) { continue }
            $status = $raw
        }
        elseif ($loc -match '^(non-?compliant|not compliant|nicht konform|error|fehler|conflict|konflikt)$') {
            $status = $loc
        }
        else { continue }
        $err = if ($errorCol -and "$($row[$errorCol])" -notin '', '0') { "$($row[$errorCol])" } else { $null }
        [ordered]@{
            name      = Get-SettingLabel "$($row[$nameCol])" "$($row[$nameCol])"
            raw       = "$($row[$nameCol])"
            state     = $status
            current   = if ($valueCol -and "$($row[$valueCol])") { "$($row[$valueCol])" } else { $null }
            errorCode = $err
            error     = $null
        }
    }
}

function ConvertTo-PolicyResult {
    param($State, [string]$DeviceId)
    $failing  = "$($State.state)" -in $failStates
    $settings = @(foreach ($s in @($State.settingStates)) {
        if (-not $s) { continue }
        if ("$($s.state)" -notin $failStates) { continue }
        [ordered]@{
            name      = Get-SettingLabel "$($s.setting)" "$($s.settingName)"
            raw       = "$($s.setting)"
            state     = "$($s.state)"
            current   = if ("$($s.currentValue)") { "$($s.currentValue)" } else { $null }
            errorCode = if ($s.errorCode -and "$($s.errorCode)" -ne '0') { "$($s.errorCode)" } else { $null }
            error     = if ("$($s.errorDescription)") { "$($s.errorDescription)" } else { $null }
        }
    })

    $note = $null
    if ($failing -and $settings.Count -eq 0) {
        if ($State.id) { $settings = @(Get-SettingFallback -DeviceId $DeviceId -PolicyId "$($State.id)") }
        if ($settings.Count -eq 0) {
            $note = 'Graph liefert für diese Richtlinie keine Einstellungsdetails. Bitte im Intune Admin Center unter Gerät > Gerätekonformität > Richtlinie prüfen.'
        }
    }

    [ordered]@{
        name     = if ($State.displayName) { "$($State.displayName)" } else { "$($State.id)" }
        state    = "$($State.state)"
        failing  = $failing
        settings = @($settings)
        note     = $note
    }
}
#endregion

#region Load devices
$select = 'id,deviceName,userPrincipalName,userDisplayName,complianceState,operatingSystem,osVersion,model,manufacturer,' +
          'serialNumber,lastSyncDateTime,enrolledDateTime,complianceGracePeriodExpirationDateTime,managedDeviceOwnerType'
$allDevices = @(Get-GraphAll -Uri "$graphBeta/deviceManagement/managedDevices?`$select=$select")
Write-Host "Managed devices found: $($allDevices.Count)" -ForegroundColor Cyan

$osPattern = @{ Windows = '^Windows'; iOS = '^(iOS|iPadOS)'; Android = '^Android'; macOS = '^macOS'; Linux = '^Linux' }
$devices = @($allDevices | Where-Object {
    $d = $_
    (-not $nameFilter -or "$($d.deviceName)" -like $nameFilter) -and
    ($osFilter.Count -eq 0 -or @($osFilter | Where-Object { "$($d.operatingSystem)" -match $osPattern[$_] }).Count -gt 0) -and
    (-not $OnlyNonCompliant -or "$($d.complianceState)" -ne 'compliant')
})
if ($devices.Count -ne $allDevices.Count) {
    Write-Host "Devices after filtering: $($devices.Count)" -ForegroundColor Cyan
}
#endregion

#region Load compliance details via batch
$detailDevices = @($devices | Where-Object { $IncludeCompliantDetails -or "$($_.complianceState)" -ne 'compliant' })
$policyStates  = @{}
$detailErrors  = @{}

if ($detailDevices.Count -gt 0) {
    Write-Host "Loading compliance details for $($detailDevices.Count) device(s) ..." -ForegroundColor Cyan
    $urlById = @{}
    foreach ($d in $detailDevices) { $urlById[$d.id] = "/deviceManagement/managedDevices/$($d.id)/deviceCompliancePolicyStates" }
    $batchFailed = $false
    try {
        $resp = Invoke-GraphBatch -RelativeUrls @($urlById.Values) -Version beta
    }
    catch {
        $msg = "Compliance details could not be loaded: $($_.Exception.Message)"
        Write-Warning $msg; $loadErrors.Add($msg)
        $resp = @{}
        $batchFailed = $true
    }
    foreach ($id in $urlById.Keys) {
        $r = $resp[$urlById[$id]]
        if ($r -and $r.status -eq 200) {
            $policyStates[$id] = @($r.body.value | Where-Object { $_ })
        }
        else {
            $detailErrors[$id] = if ($r) { "HTTP $($r.status)" } else { 'keine Antwort' }
        }
    }
    # A global batch failure is already reported above – no second, per-device message
    if ($detailErrors.Count -gt 0 -and -not $batchFailed) {
        $msg = "Compliance details not readable for $($detailErrors.Count) device(s)"
        Write-Warning $msg; $loadErrors.Add($msg)
    }
}
#endregion

#region Build result
$statusRank = @{ noncompliant = 0; grace = 1; unknown = 2; compliant = 3 }

$result = @(foreach ($d in $devices) {
    $status   = Get-StatusClass "$($d.complianceState)"
    $loaded   = $policyStates.ContainsKey($d.id)
    $policies = @()
    if ($loaded) {
        $policies = @($policyStates[$d.id] |
                      ForEach-Object { ConvertTo-PolicyResult -State $_ -DeviceId $d.id } |
                      Sort-Object { -not $_['failing'] }, { $_['name'] })
    }

    $reasons = @(foreach ($p in @($policies | Where-Object { $_['failing'] })) {
        if (@($p['settings']).Count -gt 0) {
            foreach ($s in $p['settings']) {
                $text = "$($p['name']): $($s['name'])"
                if ($s['current']) { $text += " (Ist: $($s['current']))" }
                $text
            }
        }
        else { "$($p['name']): $($p['state'])" }
    })

    [pscustomobject]@{
        DeviceName      = "$($d.deviceName)"
        User            = "$($d.userDisplayName)"
        UPN             = "$($d.userPrincipalName)"
        OS              = "$($d.operatingSystem)"
        OSVersion       = "$($d.osVersion)"
        Manufacturer    = "$($d.manufacturer)"
        Model           = "$($d.model)"
        SerialNumber    = "$($d.serialNumber)"
        Ownership       = "$($d.managedDeviceOwnerType)"
        ComplianceState = "$($d.complianceState)"
        Status          = $status
        LastSync        = ConvertTo-IsoDate $d.lastSyncDateTime
        Enrolled        = ConvertTo-IsoDate $d.enrolledDateTime
        GraceUntil      = if ($status -eq 'grace') { ConvertTo-IsoDate $d.complianceGracePeriodExpirationDateTime } else { $null }
        Reasons         = $reasons -join ' | '
        DetailsLoaded   = $loaded
        DetailsError    = $detailErrors[$d.id]
        DeviceId        = "$($d.id)"
        Policies        = $policies
    }
})
$result = @($result | Sort-Object @{ Expression = { $statusRank[$_.Status] } }, DeviceName)

$counts = [ordered]@{
    total        = $result.Count
    noncompliant = @($result | Where-Object Status -eq 'noncompliant').Count
    grace        = @($result | Where-Object Status -eq 'grace').Count
    unknown      = @($result | Where-Object Status -eq 'unknown').Count
    compliant    = @($result | Where-Object Status -eq 'compliant').Count
}
#endregion

#region HTML template (labels in German; data is injected as JSON)
$htmlTemplate = @'
<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Intune Compliance-Status</title>
<style>
:root{
  --bg:#f4f5f7;--surface:#ffffff;--text:#1c2128;--muted:#5d6672;--border:#dce0e5;--accent:#0f6cbd;
  --ok:#1a7f37;--ok-bg:#dcf7e3;--bad:#c62828;--bad-bg:#fde6e6;--warn:#8a5a00;--warn-bg:#fff1cc;--neutral:#57606a;--neutral-bg:#e9edf1;
}
@media (prefers-color-scheme: dark){
  :root{
    --bg:#111418;--surface:#1b2027;--text:#e6e8eb;--muted:#9aa4af;--border:#2e353e;--accent:#5aa2f0;
    --ok:#4ac26b;--ok-bg:#14301d;--bad:#ff7b72;--bad-bg:#3d1719;--warn:#e3b341;--warn-bg:#382a0c;--neutral:#9aa4af;--neutral-bg:#272d35;
  }
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--text);font:14px/1.45 "Segoe UI",system-ui,-apple-system,Roboto,sans-serif}
header{padding:22px 24px 6px}
h1{margin:0 0 4px;font-size:20px;font-weight:600}
.meta{color:var(--muted);font-size:12px}
.banner{margin:10px 24px 0;padding:10px 14px;border-radius:8px;background:var(--warn-bg);color:var(--warn);font-size:13px}
.banner ul{margin:6px 0 0 18px;padding:0}
.summary{display:flex;flex-wrap:wrap;gap:10px;padding:12px 24px 4px}
.stat{background:var(--surface);border:1px solid var(--border);border-top:4px solid var(--neutral);border-radius:8px;padding:8px 14px;min-width:130px;cursor:pointer;text-align:left;font:inherit;color:inherit}
.stat b{display:block;font-size:22px;font-weight:600;font-variant-numeric:tabular-nums}
.stat span{color:var(--muted);font-size:12px}
.stat.all{border-top-color:var(--accent)}
.stat.noncompliant{border-top-color:var(--bad)}
.stat.grace{border-top-color:var(--warn)}
.stat.compliant{border-top-color:var(--ok)}
.stat[aria-pressed="true"]{outline:2px solid var(--accent);outline-offset:1px}
.toolbar{padding:10px 24px 0}
.toolbar input{width:100%;max-width:440px;padding:8px 11px;border:1px solid var(--border);border-radius:6px;background:var(--surface);color:var(--text);font:inherit}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:12px;padding:14px 24px 32px}
.tile{display:flex;flex-direction:column;gap:3px;background:var(--surface);border:1px solid var(--border);border-left:5px solid var(--neutral);border-radius:8px;padding:12px 14px;cursor:pointer;text-align:left;font:inherit;color:inherit;min-width:0}
.tile:hover,.tile:focus-visible{border-color:var(--accent);outline:none}
.tile.compliant{border-left-color:var(--ok)}
.tile.noncompliant{border-left-color:var(--bad)}
.tile.grace{border-left-color:var(--warn)}
.tile .top{display:flex;justify-content:space-between;align-items:flex-start;gap:8px}
.tile .name{font-weight:600;font-size:15px;overflow-wrap:anywhere}
.tile .line{color:var(--muted);font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tile .reason{font-size:12px;color:var(--bad);margin-top:4px;overflow:hidden;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical}
.badge{flex:none;display:inline-block;padding:2px 9px;border-radius:999px;font-size:11px;font-weight:600;white-space:nowrap}
.badge.compliant{background:var(--ok-bg);color:var(--ok)}
.badge.noncompliant{background:var(--bad-bg);color:var(--bad)}
.badge.grace{background:var(--warn-bg);color:var(--warn)}
.badge.unknown{background:var(--neutral-bg);color:var(--neutral)}
.empty{grid-column:1/-1;color:var(--muted);padding:24px 0}
dialog{width:min(780px,calc(100% - 32px));padding:0;border:1px solid var(--border);border-radius:10px;background:var(--surface);color:var(--text)}
dialog::backdrop{background:rgba(0,0,0,.5)}
.dlg-head{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;padding:16px 20px;border-bottom:1px solid var(--border)}
.dlg-head h2{margin:0 0 6px;font-size:18px;overflow-wrap:anywhere}
.dlg-body{padding:16px 20px 20px;max-height:72vh;overflow:auto}
.btn{background:none;border:1px solid var(--border);border-radius:6px;color:var(--text);cursor:pointer;padding:5px 12px;font:inherit}
.btn:hover{border-color:var(--accent)}
table.info{width:100%;border-collapse:collapse;font-size:13px;margin-bottom:16px}
table.info td{padding:5px 8px;border-bottom:1px solid var(--border);vertical-align:top;overflow-wrap:anywhere}
table.info td:first-child{color:var(--muted);width:34%}
h3.sec{font-size:14px;margin:4px 0 8px}
.policy{border:1px solid var(--border);border-radius:8px;margin-bottom:10px;overflow:hidden}
.policy .p-head{display:flex;justify-content:space-between;gap:8px;padding:8px 12px;background:var(--bad-bg);color:var(--bad);font-weight:600}
.setting{padding:8px 12px;border-top:1px solid var(--border)}
.setting .s-name{font-weight:600}
.setting .s-detail{color:var(--muted);font-size:12px;overflow-wrap:anywhere}
.note{padding:8px 12px;color:var(--muted);font-size:12px}
.hint{color:var(--muted);font-size:13px;padding:10px 12px;border:1px dashed var(--border);border-radius:8px;margin-bottom:10px}
ul.others{margin:0;padding-left:18px;font-size:13px}
ul.others li span{color:var(--muted)}
a{color:var(--accent)}
@media (max-width:520px){header,.summary,.toolbar,.grid{padding-left:16px;padding-right:16px}.banner{margin-left:16px;margin-right:16px}}
</style>
</head>
<body>
<header>
  <h1>Intune Compliance-Status</h1>
  <div class="meta" id="meta"></div>
  <div class="meta" id="filter"></div>
</header>
<div id="banner"></div>
<div class="summary" id="summary"></div>
<div class="toolbar"><input id="search" type="search" placeholder="Suche nach Gerät, Benutzer, Seriennummer oder Grund …" aria-label="Suche"></div>
<main class="grid" id="grid"></main>
<dialog id="dlg" aria-labelledby="dlg-title">
  <div class="dlg-head">
    <div><h2 id="dlg-title"></h2><span id="dlg-badge"></span></div>
    <button class="btn" id="dlg-close" type="button">Schließen</button>
  </div>
  <div class="dlg-body" id="dlg-body"></div>
</dialog>
<script type="application/json" id="report-data">__REPORT_DATA__</script>
<script type="application/json" id="report-meta">__REPORT_META__</script>
<script>
(function () {
  'use strict';
  var devices = JSON.parse(document.getElementById('report-data').textContent);
  var meta = JSON.parse(document.getElementById('report-meta').textContent);

  var LABEL = { compliant: 'Compliant', noncompliant: 'Nicht compliant', grace: 'Karenzzeit', unknown: 'Unbekannt' };
  var STATE = { compliant: 'Compliant', noncompliant: 'Nicht compliant', error: 'Fehler', conflict: 'Konflikt',
                notapplicable: 'Nicht anwendbar', unknown: 'Unbekannt', ingraceperiod: 'Karenzzeit',
                notassigned: 'Nicht zugewiesen', remediated: 'Korrigiert', configmanager: 'Configuration Manager' };
  var filter = 'all';
  var query = '';

  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== undefined && text !== null) e.textContent = text;
    return e;
  }
  function stateText(s) { var k = String(s || '').toLowerCase(); return STATE[k] || s || '–'; }
  function fmtDate(v) {
    if (!v) return '–';
    var d = new Date(v);
    if (isNaN(d.getTime())) return String(v);
    return d.toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
  }
  function daysSince(v) {
    if (!v) return null;
    var d = new Date(v);
    return isNaN(d.getTime()) ? null : Math.floor((Date.now() - d.getTime()) / 86400000);
  }
  function badge(status) { return el('span', 'badge ' + status, LABEL[status] || status); }
  function failingPolicies(d) { return (d.policies || []).filter(function (p) { return p.failing; }); }
  function reasonList(d) {
    var out = [];
    failingPolicies(d).forEach(function (p) {
      var s = p.settings || [];
      if (!s.length) { if (out.indexOf(p.name) < 0) out.push(p.name); }
      s.forEach(function (x) { if (out.indexOf(x.name) < 0) out.push(x.name); });
    });
    return out;
  }
  function searchText(d) {
    return [d.name, d.user, d.upn, d.serial, d.model, d.os].concat(reasonList(d)).join(' ').toLowerCase();
  }

  /* Header, banner, summary */
  document.getElementById('meta').textContent =
    'Tenant ' + meta.tenant + ' · erstellt ' + fmtDate(meta.generated) + ' · ' + meta.account;

  var f = meta.filter || {};
  var fParts = [];
  if (f.os && f.os.length) fParts.push('Betriebssysteme: ' + f.os.join(', '));
  if (f.name) fParts.push('Gerätename: ' + f.name);
  if (f.onlyNonCompliant) fParts.push('nur nicht compliante Geräte');
  document.getElementById('filter').textContent =
    'Filter: ' + (fParts.length ? fParts.join(' · ') : 'keiner (alle Geräte)');

  if (meta.errors && meta.errors.length) {
    var b = el('div', 'banner', 'Der Report ist unvollständig – Fehler beim Laden:');
    var ul = el('ul');
    meta.errors.forEach(function (e) { ul.appendChild(el('li', null, e)); });
    b.appendChild(ul);
    document.getElementById('banner').appendChild(b);
  }

  var stats = [
    { key: 'all', label: 'Alle Geräte', n: meta.counts.total },
    { key: 'noncompliant', label: 'Nicht compliant', n: meta.counts.noncompliant },
    { key: 'grace', label: 'Karenzzeit', n: meta.counts.grace },
    { key: 'unknown', label: 'Unbekannt', n: meta.counts.unknown },
    { key: 'compliant', label: 'Compliant', n: meta.counts.compliant }
  ];
  var summary = document.getElementById('summary');
  // With "only not compliant" the compliant devices were never loaded – a "0" tile would be misleading
  if (meta.filter && meta.filter.onlyNonCompliant) {
    stats = stats.filter(function (s) { return s.key !== 'compliant'; });
    stats[0].label = 'Geräte im Report';
  }
  stats.forEach(function (s) {
    var btn = el('button', 'stat ' + s.key);
    btn.type = 'button';
    btn.setAttribute('aria-pressed', s.key === filter ? 'true' : 'false');
    btn.appendChild(el('b', null, String(s.n)));
    btn.appendChild(el('span', null, s.label));
    btn.addEventListener('click', function () {
      filter = s.key;
      Array.prototype.forEach.call(summary.children, function (c) { c.setAttribute('aria-pressed', 'false'); });
      btn.setAttribute('aria-pressed', 'true');
      render();
    });
    summary.appendChild(btn);
  });

  /* Tiles */
  var grid = document.getElementById('grid');
  function render() {
    grid.textContent = '';
    var shown = devices.filter(function (d) {
      return (filter === 'all' || d.status === filter) && (!query || searchText(d).indexOf(query) >= 0);
    });
    if (!shown.length) { grid.appendChild(el('div', 'empty', 'Keine Geräte für diese Auswahl.')); return; }
    shown.forEach(function (d) {
      var t = el('button', 'tile ' + d.status);
      t.type = 'button';
      var top = el('div', 'top');
      top.appendChild(el('div', 'name', d.name || '(ohne Namen)'));
      top.appendChild(badge(d.status));
      t.appendChild(top);
      t.appendChild(el('div', 'line', d.user || d.upn || 'Kein Benutzer'));
      t.appendChild(el('div', 'line', [d.os, d.osVersion].filter(Boolean).join(' ')));
      if (d.status !== 'compliant') {
        var r = reasonList(d);
        if (r.length) t.appendChild(el('div', 'reason', r.join(', ')));
      }
      t.addEventListener('click', function () { openDetails(d); });
      grid.appendChild(t);
    });
  }

  /* Details dialog */
  var dlg = document.getElementById('dlg');
  var dlgBody = document.getElementById('dlg-body');
  document.getElementById('dlg-close').addEventListener('click', function () { dlg.close(); });
  dlg.addEventListener('click', function (e) { if (e.target === dlg) dlg.close(); });

  function infoRow(tbl, label, value) {
    if (value === undefined || value === null || value === '') return;
    var tr = el('tr');
    tr.appendChild(el('td', null, label));
    var td = el('td');
    if (value instanceof Node) td.appendChild(value); else td.textContent = value;
    tr.appendChild(td);
    tbl.appendChild(tr);
  }

  function openDetails(d) {
    document.getElementById('dlg-title').textContent = d.name || '(ohne Namen)';
    var bw = document.getElementById('dlg-badge');
    bw.textContent = '';
    bw.appendChild(badge(d.status));
    dlgBody.textContent = '';

    var tbl = el('table', 'info');
    infoRow(tbl, 'Benutzer', d.user);
    infoRow(tbl, 'UPN', d.upn);
    infoRow(tbl, 'Betriebssystem', [d.os, d.osVersion].filter(Boolean).join(' '));
    infoRow(tbl, 'Hersteller / Modell', [d.manufacturer, d.model].filter(Boolean).join(' '));
    infoRow(tbl, 'Seriennummer', d.serial);
    infoRow(tbl, 'Besitz', d.ownership);
    var ds = daysSince(d.lastSync);
    infoRow(tbl, 'Letzter Check-in', fmtDate(d.lastSync) + (ds !== null && ds > 0 ? ' (vor ' + ds + (ds === 1 ? ' Tag)' : ' Tagen)') : ''));
    infoRow(tbl, 'Registriert', fmtDate(d.enrolled));
    if (d.graceUntil) infoRow(tbl, 'Karenzzeit bis', fmtDate(d.graceUntil));
    infoRow(tbl, 'Compliance-Status (Intune)', stateText(d.state));
    var link = el('a', null, 'Gerät im Intune Admin Center öffnen');
    link.href = 'https://intune.microsoft.com/#view/Microsoft_Intune_Devices/DeviceSettingsMenuBlade/~/overview/mdmDeviceId/' + encodeURIComponent(d.id);
    link.target = '_blank';
    link.rel = 'noopener';
    infoRow(tbl, 'Link', link);
    dlgBody.appendChild(tbl);

    if (!d.detailsLoaded) {
      if (d.detailsError) dlgBody.appendChild(el('div', 'hint', 'Compliance-Details konnten nicht geladen werden (' + d.detailsError + ').'));
      else if (d.status === 'compliant') dlgBody.appendChild(el('div', 'hint', 'Das Gerät erfüllt alle zugewiesenen Compliance-Richtlinien.'));
      dlg.showModal();
      return;
    }

    var failing = failingPolicies(d);
    if (failing.length) {
      dlgBody.appendChild(el('h3', 'sec', 'Warum nicht compliant?'));
      failing.forEach(function (p) {
        var box = el('div', 'policy');
        var head = el('div', 'p-head');
        head.appendChild(el('span', null, p.name));
        head.appendChild(el('span', null, stateText(p.state)));
        box.appendChild(head);
        (p.settings || []).forEach(function (s) {
          var row = el('div', 'setting');
          row.appendChild(el('div', 's-name', s.name));
          var parts = ['Status: ' + stateText(s.state)];
          if (s.current) parts.push('Aktueller Wert: ' + s.current);
          if (s.errorCode) parts.push('Fehlercode: ' + s.errorCode);
          if (s.error) parts.push(s.error);
          row.appendChild(el('div', 's-detail', parts.join(' · ')));
          if (s.raw && s.raw !== s.name) row.appendChild(el('div', 's-detail', s.raw));
          box.appendChild(row);
        });
        if (p.note) box.appendChild(el('div', 'note', p.note));
        dlgBody.appendChild(box);
      });
    } else if (d.status !== 'compliant') {
      dlgBody.appendChild(el('div', 'hint',
        'Intune meldet für dieses Gerät keine fehlerhafte Richtlinie. Typische Ursachen: Status noch nicht aktualisiert ' +
        '(Gerät synchronisieren), Compliance wird über Configuration Manager oder einen Partner gemeldet, oder das Gerät ' +
        'wurde noch nicht ausgewertet.'));
    }

    var others = (d.policies || []).filter(function (p) { return !p.failing; });
    if (others.length) {
      dlgBody.appendChild(el('h3', 'sec', failing.length ? 'Weitere Richtlinien' : 'Ausgewertete Richtlinien'));
      var ul = el('ul', 'others');
      others.forEach(function (p) {
        var li = el('li', null, p.name + ' ');
        li.appendChild(el('span', null, '– ' + stateText(p.state)));
        ul.appendChild(li);
      });
      dlgBody.appendChild(ul);
    }
    dlg.showModal();
  }

  document.getElementById('search').addEventListener('input', function (e) {
    query = String(e.target.value || '').trim().toLowerCase();
    render();
  });
  render();
})();
</script>
</body>
</html>
'@
#endregion

#region Output
# Data for the report (keys in lowerCamelCase for the page script)
$jsonDevices = @(foreach ($r in $result) {
    [ordered]@{
        id            = $r.DeviceId
        name          = $r.DeviceName
        user          = $r.User
        upn           = $r.UPN
        os            = $r.OS
        osVersion     = $r.OSVersion
        manufacturer  = $r.Manufacturer
        model         = $r.Model
        serial        = $r.SerialNumber
        ownership     = $r.Ownership
        state         = $r.ComplianceState
        status        = $r.Status
        lastSync      = $r.LastSync
        enrolled      = $r.Enrolled
        graceUntil    = $r.GraceUntil
        detailsLoaded = $r.DetailsLoaded
        detailsError  = $r.DetailsError
        policies      = @($r.Policies)
    }
})
$metaObj = [ordered]@{
    tenant    = "$($ctx.TenantId)"
    account   = "$identity"
    generated = (Get-Date).ToUniversalTime().ToString('o')
    counts    = $counts
    errors    = @($loadErrors)
    filter    = [ordered]@{
        name             = $nameFilter
        os               = @($osFilter)
        onlyNonCompliant = [bool]$OnlyNonCompliant
    }
}

# "<" is escaped so that no value can close the <script> block
$dataJson = (ConvertTo-Json -InputObject $jsonDevices -Depth 10 -Compress).Replace('<', '\u003c')
$metaJson = (ConvertTo-Json -InputObject $metaObj -Depth 5 -Compress).Replace('<', '\u003c')
$html     = $htmlTemplate.Replace('__REPORT_DATA__', $dataJson).Replace('__REPORT_META__', $metaJson)

$OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
$outDir = Split-Path -Path $OutputPath -Parent
if ($outDir -and -not (Test-Path -Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
[IO.File]::WriteAllText($OutputPath, $html, [Text.UTF8Encoding]::new($false))

Write-Host ""
Write-Host "Compliance summary" -ForegroundColor Cyan
Write-Host "  Devices:        $($counts.total)"
Write-Host "  Not compliant:  $($counts.noncompliant)" -ForegroundColor $(if ($counts.noncompliant) { 'Red' } else { 'Gray' })
Write-Host "  Grace period:   $($counts.grace)"        -ForegroundColor $(if ($counts.grace) { 'Yellow' } else { 'Gray' })
Write-Host "  Unknown:        $($counts.unknown)"
Write-Host "  Compliant:      $($counts.compliant)"    -ForegroundColor Green

$problems = @($result | Where-Object { $_.Status -in 'noncompliant', 'grace' })
if ($problems.Count -gt 0) {
    $problems | Format-Table DeviceName, User, Status, Reasons -AutoSize -Wrap | Out-Host
}

if ($CsvPath) {
    $CsvPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($CsvPath)
    $csvDir = Split-Path -Path $CsvPath -Parent
    if ($csvDir -and -not (Test-Path -Path $csvDir)) { New-Item -ItemType Directory -Path $csvDir -Force | Out-Null }
    $encoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
    $result | Select-Object DeviceName, User, UPN, OS, OSVersion, Manufacturer, Model, SerialNumber, ComplianceState, Status,
                            LastSync, GraceUntil, Reasons, DeviceId |
        Export-Csv -Path $CsvPath -Delimiter $Delimiter -NoTypeInformation -Encoding $encoding
}

if ($loadErrors.Count -gt 0) {
    Write-Warning "The report is INCOMPLETE – $($loadErrors.Count) error(s) while loading:"
    $loadErrors | ForEach-Object { Write-Warning "  $_" }
}

# Saved files last, so they are not scrolled away by the table or warnings
Write-Host ""
Write-Host "Report saved" -ForegroundColor Cyan
Write-Host "  HTML: $OutputPath" -ForegroundColor Green
if ($CsvPath) {
    Write-Host "  CSV:  $CsvPath ($($result.Count) rows)" -ForegroundColor Green
}
if ($Open) {
    Invoke-Item -Path $OutputPath
}
else {
    # Single quotes inside a single-quoted string must be doubled, so the line stays copy-paste-safe
    Write-Host "  Open: Invoke-Item '$($OutputPath.Replace("'", "''"))'   (or run the script with -Open)" -ForegroundColor Gray
}
if ($PassThru) { $result }
#endregion
