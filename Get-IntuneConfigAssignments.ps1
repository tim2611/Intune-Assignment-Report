<#
.SYNOPSIS
    Listet alle Intune-Gerätekonfigurationen inkl. Zuweisungen (Gruppe, Include/Exclude, Filter).

.DESCRIPTION
    Erfasst folgende Profiltypen über Microsoft Graph (beta):
      - Klassische Templates                         /deviceManagement/deviceConfigurations
      - Settings Catalog / Endpoint Security (neu)   /deviceManagement/configurationPolicies
      - Administrative Templates (ADMX)              /deviceManagement/groupPolicyConfigurations
      - Endpoint Security (Legacy-Intents)           /deviceManagement/intents

    Fehler beim Laden einzelner Profiltypen brechen den Lauf nicht ab; sie werden als Warnung
    gemeldet und am Ende zusammengefasst. Gruppennamen und Intent-Zuweisungen werden per
    JSON-Batch ($batch, 20 Requests pro Aufruf, Retry bei HTTP 429) aufgelöst.

    Benötigte Berechtigungen (delegiert oder Application):
      DeviceManagementConfiguration.Read.All (oder ReadWrite.All)
      Group.Read.All (oder Group.ReadWrite.All / Directory.Read.All)

    Anmeldung:
      Ohne bestehende Verbindung meldet sich das Skript interaktiv (delegiert) an.
      Für App-Only (z. B. Automatisierung) vorher selbst verbinden, z. B.:
        Connect-MgGraph -TenantId <Tenant> -ClientId <AppId> -CertificateThumbprint <Thumbprint>
      Das Skript prüft die Application Permissions dann per Test-Request.

.PARAMETER TenantId
    Optional: Tenant-ID oder primäre Domain. Erzwingt die Verbindung mit genau diesem Tenant
    (wichtig bei Multi-Tenant-Admins / Gastkonten).

    Ausgabeformat (Standard): EINE Zeile pro Richtlinie. Alle Zuweisungen stehen in der Spalte
    "Gruppe", getrennt durch " | ", Include vor Exclude, z. B.:
      GRP-Win-Pilot | GRP-Win-Prod [Filter: Nur Laptops (include)] | Exclude: GRP-Win-Kiosk
    Filter stehen dabei inline hinter der jeweiligen Gruppe. Die Spalte "Zuweisungen" zählt alle
    Zuweisungen der Richtlinie (Include, Exclude, tenantweit, unbekannter Zieltyp).
    Mit -OneRowPerAssignment gibt es stattdessen eine Zeile pro Zuweisung mit eigenen Spalten
    für Modus, Filter und FilterModus.

.PARAMETER GroupName
    Optional: Beantwortet die Frage "Welche Richtlinien sind dieser Gruppe zugewiesen?".
    Ausgegeben werden Richtlinien mit mindestens einer direkten Zuweisung an eine Gruppe, deren Name
    dem Muster entspricht (Wildcards erlaubt; Include oder Exclude). Von diesen Richtlinien werden
    immer ALLE Zuweisungen angezeigt (auch andere Gruppen, Excludes und tenantweite Zuweisungen).
    Verschachtelte Mitgliedschaften werden nicht aufgelöst. Richtlinien ohne Zuweisung erscheinen
    mit -GroupName grundsätzlich nicht.

.PARAMETER IncludeTenantWide
    Nur mit -GroupName: zusätzlich alle Richtlinien ausgeben, die "Alle Geräte" oder "Alle Benutzer"
    zugewiesen sind – auch ohne Treffer auf die Gruppe. Damit enthält der Report alles, was Mitglieder
    der Gruppe über direkte UND tenantweite Zuweisungen erreichen kann (Excludes und Filter beachten).

.PARAMETER IncludeUnassigned
    Auch Richtlinien ohne Zuweisung ausgeben. Hat mit -GroupName keine Wirkung.

.PARAMETER ExportPath
    Pfad der CSV-Datei (Semikolon-getrennt, UTF-8 mit BOM für Excel). Zielordner wird bei Bedarf angelegt.
    Standard: %TEMP%\Intune-ConfigAssignments_<Zeitstempel>.csv

.PARAMETER OneRowPerAssignment
    Eine Zeile pro Zuweisung statt pro Richtlinie (mit Spalten Modus, Filter, FilterModus) –
    sinnvoll für Pivot-Auswertungen in Excel.

.PARAMETER GridView
    Ergebnis zusätzlich in Out-GridView anzeigen.

.PARAMETER PassThru
    Ergebnisobjekte in die Pipeline ausgeben (statt nur Tabelle/CSV).

.EXAMPLE
    .\Get-IntuneConfigAssignments.ps1 -TenantId contoso.onmicrosoft.com
.EXAMPLE
    .\Get-IntuneConfigAssignments.ps1 -GroupName "GRP-Intune-Win-*" -IncludeTenantWide -GridView
.EXAMPLE
    .\Get-IntuneConfigAssignments.ps1 -IncludeUnassigned -ExportPath C:\Reports\assignments.csv
.EXAMPLE
    $r = .\Get-IntuneConfigAssignments.ps1 -PassThru; $r | Where-Object Zuweisungen -gt 1
.EXAMPLE
    .\Get-IntuneConfigAssignments.ps1 -OneRowPerAssignment

.NOTES
    Autor:    Timo Schlappinger
    Version:  1.0.0
    Projekt:  https://github.com/tim2611/Intune-Assignment-Report
    Benötigt: Modul Microsoft.Graph.Authentication, PowerShell 5.1 oder 7.x
#>
[CmdletBinding()]
[OutputType([pscustomobject])]
param(
    [string]$TenantId,
    [string]$GroupName,
    [switch]$IncludeTenantWide,
    [switch]$IncludeUnassigned,
    [string]$ExportPath = (Join-Path $env:TEMP "Intune-ConfigAssignments_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"),
    [switch]$OneRowPerAssignment,
    [switch]$GridView,
    [switch]$PassThru
)

$ErrorActionPreference = 'Stop'
$graphBeta  = 'https://graph.microsoft.com/beta'
$loadErrors = [System.Collections.Generic.List[string]]::new()

#region Verbindung & Berechtigungen
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    throw "Modul fehlt: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
}
Import-Module Microsoft.Graph.Authentication

# Benötigte Berechtigung -> akzeptierte (gleich- oder höherwertige) Berechtigungen
$scopeMap = [ordered]@{
    'DeviceManagementConfiguration.Read.All' = @('DeviceManagementConfiguration.Read.All', 'DeviceManagementConfiguration.ReadWrite.All')
    'Group.Read.All'                         = @('Group.Read.All', 'Group.ReadWrite.All', 'Directory.Read.All', 'Directory.ReadWrite.All')
}

function Get-MissingScopes {
    $ctx = Get-MgContext
    if (-not $ctx) { return @($scopeMap.Keys) }
    # Normalisierung: falls Scopes als volle URL geliefert werden, nur den Kurznamen verwenden
    $have = @($ctx.Scopes | ForEach-Object { ($_ -split '/')[-1] })
    $missing = foreach ($required in $scopeMap.Keys) {
        if (-not ($scopeMap[$required] | Where-Object { $_ -in $have })) { $required }
    }
    return @($missing)
}

$ctx         = Get-MgContext
$isAppOnly   = $ctx -and ("$($ctx.AuthType)" -eq 'AppOnly')
$wrongTenant = $ctx -and $TenantId -and ($ctx.TenantId -ne $TenantId)

if ($isAppOnly) {
    if ($wrongTenant) {
        throw "Bestehende App-Verbindung zeigt auf Tenant $($ctx.TenantId), nicht auf $TenantId. Bitte Disconnect-MgGraph und mit der passenden App-Registrierung neu verbinden."
    }
}
elseif (-not $ctx -or $wrongTenant -or (Get-MissingScopes).Count -gt 0) {
    $connectParams = @{ Scopes = @($scopeMap.Keys); NoWelcome = $true }
    if ($TenantId) { $connectParams.TenantId = $TenantId }
    Connect-MgGraph @connectParams
}

$ctx = Get-MgContext
if ("$($ctx.AuthType)" -eq 'AppOnly') {
    # Application Permissions sind im Context nicht zuverlässig sichtbar -> Funktionstest per Probe-Request
    $probes = [ordered]@{
        'DeviceManagementConfiguration.Read.All' = "$graphBeta/deviceManagement/deviceConfigurations?`$top=1&`$select=id"
        'Group.Read.All'                         = 'https://graph.microsoft.com/v1.0/groups?$top=1&$select=id'
    }
    $failed = foreach ($perm in $probes.Keys) {
        try   { Invoke-MgGraphRequest -Method GET -Uri $probes[$perm] | Out-Null }
        catch { $perm }
    }
    if ($failed) {
        throw "App-Registrierung fehlt (vermutlich) die Application Permission: $(@($failed) -join ', ') (inkl. Admin Consent)"
    }
}
else {
    $missing = Get-MissingScopes
    if ($missing.Count -gt 0) {
        throw "Fehlende delegierte Berechtigungen: $($missing -join ', ')"
    }
}
$identity = if ($ctx.Account) { $ctx.Account } else { $ctx.AppName }
Write-Host "Verbunden mit Tenant $($ctx.TenantId) als $identity ($($ctx.AuthType))" -ForegroundColor Cyan
#endregion

#region Hilfsfunktionen
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

function Get-GraphAllSafe {
    param([Parameter(Mandatory)][string]$Uri, [Parameter(Mandatory)][string]$Label)
    try {
        Get-GraphAll -Uri $Uri
    }
    catch {
        $msg = "$Label konnte nicht geladen werden: $($_.Exception.Message)"
        Write-Warning $msg
        $script:loadErrors.Add($msg)
    }
}

function Invoke-GraphBatch {
    <#
      Führt GET-Requests per JSON-Batch aus (max. 20 pro Aufruf).
      Einzelne 429/503/504-Antworten im Batch werden bis zu 3x mit Retry-After erneut versucht,
      ebenso 5xx-/Transportfehler des Batch-POST selbst. Status 0 = Batch-POST nie erfolgreich.
      Rückgabe: Hashtable relativeUrl -> Antwortobjekt (status, body).
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

            # Batch-POST selbst absichern: 5xx/Transportfehler -> Retry, 4xx (z. B. 401/403) -> sofort abbrechen
            try {
                $resp = Invoke-MgGraphRequest -Method POST -Uri $batchUri -Body $body -ContentType 'application/json'
            }
            catch {
                $code = $null
                if ($_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode }
                if ($attempt -eq $maxAttempts -or ($code -and $code -lt 500 -and $code -ne 429)) { throw }
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

        # Nach allen Versuchen weiterhin offene Requests mit ihrer letzten Antwort (Status) übernehmen
        foreach ($url in $pending) {
            $results[$url] = if ($lastResp.ContainsKey($url)) { $lastResp[$url] } else { @{ status = 0; body = $null } }
        }
    }
    return $results
}

function Get-NormalizedPlatform {
    param([string]$Value)
    if (-not $Value) { return 'Unbekannt' }
    $mapped = foreach ($part in ($Value -split '\s*,\s*')) {
        switch -Regex ($part) {
            '^(windows|editionUpgrade|sharedPC)' { 'Windows'; break }
            '^ios'                               { 'iOS';     break }
            '^macos'                             { 'macOS';   break }
            '^(android|aosp)'                    { 'Android'; break }
            '^linux'                             { 'Linux';   break }
            default                              { $part }
        }
    }
    return (($mapped | Select-Object -Unique) -join ', ')
}
#endregion

#region Assignment Filters laden
$filterNames = @{}
foreach ($f in (Get-GraphAllSafe -Uri "$graphBeta/deviceManagement/assignmentFilters" -Label 'Assignment Filters')) {
    $filterNames[$f.id] = $f.displayName
}
#endregion

#region Profile laden
$profiles = [System.Collections.Generic.List[object]]::new()

# Klassische Templates
foreach ($p in (Get-GraphAllSafe -Uri "$graphBeta/deviceManagement/deviceConfigurations?`$expand=assignments" -Label 'Templates (deviceConfigurations)')) {
    $odata = $p.'@odata.type' -replace '^#microsoft\.graph\.', ''
    $profiles.Add([pscustomobject]@{
        Typ         = 'Template'
        Detailtyp   = $odata
        Name        = $p.displayName
        Plattform   = Get-NormalizedPlatform $odata
        Id          = $p.id
        Assignments = @($p.assignments)
    })
}

# Settings Catalog + neue Endpoint-Security-Richtlinien + neue Security Baselines
foreach ($p in (Get-GraphAllSafe -Uri "$graphBeta/deviceManagement/configurationPolicies?`$expand=assignments" -Label 'Settings Catalog (configurationPolicies)')) {
    $family = $p.templateReference.templateFamily
    $typ = if ($family -and $family -ne 'none') { "Endpoint Security / Baseline" } else { 'Settings Catalog' }
    $profiles.Add([pscustomobject]@{
        Typ         = $typ
        Detailtyp   = if ($family -and $family -ne 'none') { $family } else { @($p.technologies) -join ', ' }
        Name        = $p.name
        Plattform   = Get-NormalizedPlatform $p.platforms
        Id          = $p.id
        Assignments = @($p.assignments)
    })
}

# Administrative Templates (ADMX)
foreach ($p in (Get-GraphAllSafe -Uri "$graphBeta/deviceManagement/groupPolicyConfigurations?`$expand=assignments" -Label 'Administrative Templates (groupPolicyConfigurations)')) {
    $profiles.Add([pscustomobject]@{
        Typ         = 'Administrative Template'
        Detailtyp   = 'groupPolicyConfiguration'
        Name        = $p.displayName
        Plattform   = 'Windows'
        Id          = $p.id
        Assignments = @($p.assignments)
    })
}

# Endpoint Security Legacy-Intents (kein $expand möglich -> Zuweisungen per Batch)
$intents = @(Get-GraphAllSafe -Uri "$graphBeta/deviceManagement/intents" -Label 'Endpoint Security Legacy-Intents')
if ($intents.Count -gt 0) {
    $intentUrls = @($intents | ForEach-Object { "/deviceManagement/intents/$($_.id)/assignments" })
    try {
        $intentAssignments = Invoke-GraphBatch -RelativeUrls $intentUrls -Version beta
    }
    catch {
        $msg = "Zuweisungen der Legacy-Intents konnten nicht geladen werden: $($_.Exception.Message)"
        Write-Warning $msg; $loadErrors.Add($msg)
        $intentAssignments = @{}
    }
    foreach ($p in $intents) {
        $r = $intentAssignments["/deviceManagement/intents/$($p.id)/assignments"]
        if ($r -and $r.status -ne 200) {
            $msg = "Intent '$($p.displayName)': Zuweisungen nicht lesbar (HTTP $($r.status))"
            Write-Warning $msg; $loadErrors.Add($msg)
        }
        $profiles.Add([pscustomobject]@{
            Typ         = 'Endpoint Security (Legacy-Intent)'
            Detailtyp   = $p.templateId
            Name        = $p.displayName
            Plattform   = 'Windows'
            Id          = $p.id
            Assignments = if ($r -and $r.status -eq 200) { @($r.body.value) } else { @() }
        })
    }
}

Write-Host "Profile gefunden: $($profiles.Count)" -ForegroundColor Cyan
#endregion

#region Zuweisungen aufbereiten
$emptyGuid = '00000000-0000-0000-0000-000000000000'

$rows = foreach ($p in $profiles) {
    $assignments = @($p.Assignments | Where-Object { $_ })

    if ($assignments.Count -eq 0) {
        [pscustomobject]@{
            Richtlinie = $p.Name; Typ = $p.Typ; Plattform = $p.Plattform; Detailtyp = $p.Detailtyp
            Modus = '-'; Bereich = 'Keine'; Gruppe = '<nicht zugewiesen>'
            Filter = $null; FilterModus = $null; GruppenId = $null; RichtlinienId = $p.Id
        }
        continue
    }

    foreach ($a in $assignments) {
        $t = $a.target
        $groupId = $null; $gruppe = $null; $bereich = 'Gruppe'

        switch ($t.'@odata.type') {
            '#microsoft.graph.groupAssignmentTarget'            { $modus = 'Include'; $groupId = $t.groupId }
            '#microsoft.graph.exclusionGroupAssignmentTarget'   { $modus = 'Exclude'; $groupId = $t.groupId }
            '#microsoft.graph.allDevicesAssignmentTarget'       { $modus = 'Include'; $gruppe = 'Alle Geräte';   $bereich = 'Tenantweit' }
            '#microsoft.graph.allLicensedUsersAssignmentTarget' { $modus = 'Include'; $gruppe = 'Alle Benutzer'; $bereich = 'Tenantweit' }
            default                                             { $modus = 'Unbekannt'; $gruppe = $t.'@odata.type'; $bereich = 'Unbekannt' }
        }

        $filterId  = $t.deviceAndAppManagementAssignmentFilterId
        $hasFilter = $filterId -and $filterId -ne $emptyGuid
        $filter    = if (-not $hasFilter) { $null }
                     elseif ($filterNames.ContainsKey($filterId)) { $filterNames[$filterId] }
                     else { "<Filter gelöscht/unbekannt: $filterId>" }

        [pscustomobject]@{
            Richtlinie    = $p.Name
            Typ           = $p.Typ
            Plattform     = $p.Plattform
            Detailtyp     = $p.Detailtyp
            Modus         = $modus
            Bereich       = $bereich
            Gruppe        = $gruppe
            Filter        = $filter
            FilterModus   = if ($hasFilter) { $t.deviceAndAppManagementAssignmentFilterType } else { $null }
            GruppenId     = $groupId
            RichtlinienId = $p.Id
        }
    }
}
$rows = @($rows)
#endregion

#region Gruppennamen per Batch auflösen
$groupIds = @($rows | Where-Object { $_.GruppenId } | Select-Object -ExpandProperty GruppenId -Unique)
if ($groupIds.Count -gt 0) {
    $groupUrls = @($groupIds | ForEach-Object { "/groups/$_`?`$select=id,displayName" })
    try {
        $groupResp = Invoke-GraphBatch -RelativeUrls $groupUrls -Version v1.0
    }
    catch {
        $msg = "Gruppennamen konnten nicht aufgelöst werden: $($_.Exception.Message)"
        Write-Warning $msg; $loadErrors.Add($msg)
        $groupResp = @{}
    }
    $groupNames = @{}
    foreach ($id in $groupIds) {
        $r = $groupResp["/groups/$id`?`$select=id,displayName"]
        $groupNames[$id] = if     (-not $r)            { "<nicht aufgelöst: $id>" }
                           elseif ($r.status -eq 200)  { $r.body.displayName }
                           elseif ($r.status -eq 404)  { "<Gruppe gelöscht: $id>" }
                           else                        { "<Fehler HTTP $($r.status): $id>" }
    }
    foreach ($row in $rows) {
        if ($row.GruppenId) { $row.Gruppe = $groupNames[$row.GruppenId] }
    }
}
#endregion

#region Filtern
$filtered = $rows
if ($GroupName) {
    # Richtlinien ermitteln, die mindestens eine passende Zuweisung haben ...
    $matchingIds = @($rows | Where-Object {
        ($_.Bereich -eq 'Gruppe' -and $_.Gruppe -like $GroupName) -or
        ($IncludeTenantWide -and $_.Bereich -eq 'Tenantweit')
    } | Select-Object -ExpandProperty RichtlinienId -Unique)
    # ... und von diesen ALLE Zuweisungen behalten (Gesamtbild inkl. Excludes)
    $filtered = $rows | Where-Object { $_.RichtlinienId -in $matchingIds }
}
elseif (-not $IncludeUnassigned) {
    $filtered = $rows | Where-Object { $_.Bereich -ne 'Keine' }
}
$filtered = @($filtered)

$modusOrder = @{ Expression = { switch ($_.Modus) { 'Include' { 0 } 'Exclude' { 1 } default { 2 } } } }
#endregion

#region Ergebnis aufbauen
if ($OneRowPerAssignment) {
    # Eine Zeile pro Zuweisung (Rohdaten, z. B. für Pivot-Auswertungen)
    $result = @($filtered | Sort-Object Richtlinie, RichtlinienId, $modusOrder, Gruppe)
}
else {
    # Eine Zeile pro Richtlinie, alle Zuweisungen in der Spalte "Gruppe" zusammengefasst
    $result = @($filtered | Group-Object RichtlinienId | ForEach-Object {
        $items = @($_.Group | Sort-Object $modusOrder, Gruppe)
        $first = $items[0]

        $labels = foreach ($i in $items) {
            $label = if ($i.Modus -eq 'Exclude') { "Exclude: $($i.Gruppe)" } else { $i.Gruppe }
            if ($i.Filter) { $label += " [Filter: $($i.Filter) ($($i.FilterModus))]" }
            $label
        }

        [pscustomobject]@{
            Richtlinie    = $first.Richtlinie
            Typ           = $first.Typ
            Plattform     = $first.Plattform
            Detailtyp     = $first.Detailtyp
            Zuweisungen   = @($items | Where-Object { $_.Bereich -ne 'Keine' }).Count
            Gruppe        = $labels -join ' | '
            GruppenIds    = @($items | Where-Object { $_.GruppenId } | ForEach-Object { $_.GruppenId }) -join ' | '
            RichtlinienId = $first.RichtlinienId
        }
    } | Sort-Object Richtlinie, RichtlinienId)
}
#endregion

#region Ausgabe
if ($result.Count -eq 0) {
    Write-Warning "Keine Zuweisungen gefunden, die den Kriterien entsprechen."
}
else {
    if ($OneRowPerAssignment) {
        $result | Format-Table Richtlinie, Typ, Plattform, Modus, Gruppe, Filter, FilterModus -AutoSize -Wrap | Out-Host
    }
    else {
        $result | Format-Table Richtlinie, Typ, Plattform, Zuweisungen, Gruppe -AutoSize -Wrap | Out-Host
    }

    $exportDir = Split-Path -Path $ExportPath -Parent
    if ($exportDir -and -not (Test-Path -Path $exportDir)) {
        New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
    }
    $encoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
    $result | Export-Csv -Path $ExportPath -Delimiter ';' -NoTypeInformation -Encoding $encoding
    Write-Host "CSV exportiert: $ExportPath ($($result.Count) Zeilen)" -ForegroundColor Green

    if ($GridView) { $result | Out-GridView -Title 'Intune Konfigurationsprofile – Zuweisungen' }
}

if ($loadErrors.Count -gt 0) {
    Write-Warning "Der Report ist UNVOLLSTÄNDIG – $($loadErrors.Count) Fehler beim Laden:"
    $loadErrors | ForEach-Object { Write-Warning "  $_" }
}

if ($PassThru) { $result }
#endregion
