<#
.SYNOPSIS
    Listet alle Intune-Compliance-Richtlinien inkl. Zuweisungen (Gruppe, Include/Exclude, Filter)
    und Aktionen bei Nichtkonformität.

.DESCRIPTION
    Erfasst folgende Richtlinientypen über Microsoft Graph (beta):
      - Compliance-Richtlinien (klassisch)          /deviceManagement/deviceCompliancePolicies
      - Compliance-Richtlinien (Settings Catalog)   /deviceManagement/compliancePolicies   (z. B. Linux)

    Pro Richtlinie werden zusätzlich die Aktionen bei Nichtkonformität ausgegeben
    (z. B. als nicht konform markieren, E-Mail, Gerät sperren, Außerbetriebnahme) inkl. Zeitpunkt,
    Benachrichtigungsvorlage und CC-Gruppen. Außerdem zeigt das Skript die tenantweiten
    Compliance-Einstellungen an ("Geräte ohne zugewiesene Compliance-Richtlinie als ..." und
    Gültigkeitszeitraum des Compliance-Status).

    Fehler beim Laden einzelner Richtlinientypen brechen den Lauf nicht ab; sie werden als Warnung
    gemeldet und am Ende zusammengefasst. Gruppennamen und Aktionen werden per JSON-Batch
    ($batch, 20 Requests pro Aufruf, Retry bei HTTP 429/503/504) aufgelöst.

    Benötigte Berechtigungen (delegiert oder Application):
      DeviceManagementConfiguration.Read.All (oder ReadWrite.All)
      Group.Read.All (oder Group.ReadWrite.All / Directory.Read.All)
    Optional (nur für Namen der Benachrichtigungsvorlagen und Tenant-Einstellungen):
      DeviceManagementServiceConfig.Read.All – fehlt sie, werden Vorlagen-IDs angezeigt.

    Anmeldung:
      Ohne bestehende Verbindung meldet sich das Skript interaktiv (delegiert) an.
      Für App-Only (z. B. Automatisierung) vorher selbst verbinden, z. B.:
        Connect-MgGraph -TenantId <Tenant> -ClientId <AppId> -CertificateThumbprint <Thumbprint>
      Das Skript prüft die Application Permissions dann per Test-Request.

    Ausgabeformat (Standard): EINE Zeile pro Richtlinie. Alle Zuweisungen stehen in der Spalte
    "Gruppe", getrennt durch " | ", Include vor Exclude, Filter inline, z. B.:
      GRP-Win-Pilot | GRP-Win-Prod [Filter: Nur Laptops (include)] | Exclude: GRP-Win-Kiosk
    Die Spalte "Aktionen" enthält alle Aktionen bei Nichtkonformität, chronologisch, z. B.:
      Als nicht konform markieren: sofort | E-Mail an Endbenutzer: nach 1 Tag (Vorlage: Hinweis) | ...
    Die Spalte "NichtKonformNach" zeigt separat, wann ein Gerät als nicht konform markiert wird.
    Mit -OneRowPerAssignment gibt es stattdessen eine Zeile pro Zuweisung mit eigenen Spalten
    für Modus, Filter und FilterModus.

.PARAMETER TenantId
    Optional: Tenant-ID oder primäre Domain. Erzwingt die Verbindung mit genau diesem Tenant
    (wichtig bei Multi-Tenant-Admins / Gastkonten).

.PARAMETER GroupName
    Optional: Beantwortet die Frage "Welche Compliance-Richtlinien sind dieser Gruppe zugewiesen?".
    Ausgegeben werden Richtlinien mit mindestens einer direkten Zuweisung an eine Gruppe, deren Name
    dem Muster entspricht (Wildcards erlaubt; Include oder Exclude). Von diesen Richtlinien werden
    immer ALLE Zuweisungen angezeigt (auch andere Gruppen, Excludes und tenantweite Zuweisungen).
    Verschachtelte Mitgliedschaften werden nicht aufgelöst. Richtlinien ohne Zuweisung erscheinen
    mit -GroupName grundsätzlich nicht.

.PARAMETER IncludeTenantWide
    Nur mit -GroupName: zusätzlich alle Richtlinien ausgeben, die "Alle Geräte" oder "Alle Benutzer"
    zugewiesen sind – auch ohne Treffer auf die Gruppe.

.PARAMETER IncludeUnassigned
    Auch Richtlinien ohne Zuweisung ausgeben. Hat mit -GroupName keine Wirkung.

.PARAMETER OneRowPerAssignment
    Eine Zeile pro Zuweisung statt pro Richtlinie (mit Spalten Modus, Filter, FilterModus) –
    sinnvoll für Pivot-Auswertungen in Excel.

.PARAMETER ExportPath
    Pfad der CSV-Datei (Semikolon-getrennt, UTF-8 mit BOM für Excel). Zielordner wird bei Bedarf angelegt.
    Standard: %TEMP%\Intune-ComplianceAssignments_<Zeitstempel>.csv

.PARAMETER GridView
    Ergebnis zusätzlich in Out-GridView anzeigen.

.PARAMETER PassThru
    Ergebnisobjekte in die Pipeline ausgeben (statt nur Tabelle/CSV).

.EXAMPLE
    .\Get-IntuneComplianceAssignments.ps1 -TenantId contoso.onmicrosoft.com
.EXAMPLE
    .\Get-IntuneComplianceAssignments.ps1 -GroupName "GRP-Intune-Win-*" -IncludeTenantWide -GridView
.EXAMPLE
    .\Get-IntuneComplianceAssignments.ps1 -IncludeUnassigned -ExportPath C:\Reports\compliance.csv
.EXAMPLE
    $r = .\Get-IntuneComplianceAssignments.ps1 -PassThru; $r | Where-Object NichtKonformNach -ne 'sofort'
.EXAMPLE
    .\Get-IntuneComplianceAssignments.ps1 -OneRowPerAssignment

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
    [string]$ExportPath = (Join-Path $env:TEMP "Intune-ComplianceAssignments_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"),
    [switch]$OneRowPerAssignment,
    [switch]$GridView,
    [switch]$PassThru
)

$ErrorActionPreference = 'Stop'
$graphBeta  = 'https://graph.microsoft.com/beta'
$emptyGuid  = '00000000-0000-0000-0000-000000000000'
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
        'DeviceManagementConfiguration.Read.All' = "$graphBeta/deviceManagement/deviceCompliancePolicies?`$top=1&`$select=id"
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
            '^windows'        { 'Windows'; break }
            '^ios'            { 'iOS';     break }
            '^macos'          { 'macOS';   break }
            '^(android|aosp)' { 'Android'; break }
            '^linux'          { 'Linux';   break }
            default           { $part }
        }
    }
    return (($mapped | Select-Object -Unique) -join ', ')
}

function Format-GracePeriod {
    # Standard: "nach 3 Tagen" (für die Aktionsliste); -Short: "3 Tage" (für die Spalte NichtKonformNach)
    param($Hours, [switch]$Short)
    $h = [int]$Hours
    if ($h -le 0) { return 'sofort' }
    if ($h % 24 -ne 0) {
        if ($Short) { return "$h Std." } else { return "nach $h Std." }
    }
    $d = $h / 24
    if ($Short) { if ($d -eq 1) { return '1 Tag' } else { return "$d Tage" } }
    if ($d -eq 1) { return 'nach 1 Tag' }
    return "nach $d Tagen"
}

# Anzeigenamen der Aktionen (wie im Intune-Portal)
$actionLabels = @{
    'block'                        = 'Als nicht konform markieren'
    'notification'                 = 'E-Mail an Endbenutzer'
    'pushNotification'             = 'Push-Benachrichtigung an Endbenutzer'
    'remoteLock'                   = 'Gerät remote sperren'
    'retire'                       = 'Gerät zur Außerbetriebnahme-Liste hinzufügen'
    'wipe'                         = 'Gerät zurücksetzen (Wipe)'
    'removeResourceAccessProfiles' = 'Ressourcenzugriffsprofile entfernen'
    'noAction'                     = 'Keine Aktion'
}
$actionOrder = @('block', 'notification', 'pushNotification', 'remoteLock', 'retire', 'wipe', 'removeResourceAccessProfiles', 'noAction')
#endregion

#region Tenant-Einstellungen, Filter und Benachrichtigungsvorlagen laden
try {
    $dm = Invoke-MgGraphRequest -Method GET -Uri "$graphBeta/deviceManagement?`$select=settings"
    $noPolicyState = if ($dm.settings.secureByDefault) { 'Nicht konform' } else { 'Konform' }
    Write-Host ""
    Write-Host "Tenantweite Compliance-Einstellungen" -ForegroundColor Cyan
    Write-Host "  Geräte ohne zugewiesene Compliance-Richtlinie gelten als: $noPolicyState"
    Write-Host "  Gültigkeitszeitraum des Compliance-Status:                $($dm.settings.deviceComplianceCheckinThresholdDays) Tage"
    Write-Host ""
}
catch {
    Write-Warning "Tenantweite Compliance-Einstellungen nicht lesbar ($($_.Exception.Message)). Benötigt ggf. DeviceManagementServiceConfig.Read.All."
}

$filterNames = @{}
foreach ($f in (Get-GraphAllSafe -Uri "$graphBeta/deviceManagement/assignmentFilters" -Label 'Assignment Filters')) {
    $filterNames[$f.id] = $f.displayName
}

# Vorlagennamen sind optional (eigene Berechtigung) – Fehler machen den Report nicht unvollständig
$templateNames = @{}
try {
    foreach ($t in (Get-GraphAll -Uri "$graphBeta/deviceManagement/notificationMessageTemplates?`$select=id,displayName")) {
        $templateNames[$t.id] = $t.displayName
    }
}
catch {
    Write-Warning "Namen der Benachrichtigungsvorlagen nicht lesbar – es werden Vorlagen-IDs angezeigt. Benötigt ggf. DeviceManagementServiceConfig.Read.All."
}
#endregion

#region Richtlinien laden
$profiles = [System.Collections.Generic.List[object]]::new()

# Klassische Compliance-Richtlinien
foreach ($p in (Get-GraphAllSafe -Uri "$graphBeta/deviceManagement/deviceCompliancePolicies?`$expand=assignments" -Label 'Compliance-Richtlinien (deviceCompliancePolicies)')) {
    $odata = $p.'@odata.type' -replace '^#microsoft\.graph\.', ''
    $profiles.Add([pscustomobject]@{
        Typ         = 'Compliance'
        Detailtyp   = $odata
        Name        = $p.displayName
        Plattform   = Get-NormalizedPlatform $odata
        Id          = $p.id
        Assignments = @($p.assignments)
        ActionsUrl  = "/deviceManagement/deviceCompliancePolicies/$($p.id)/scheduledActionsForRule?`$expand=scheduledActionConfigurations"
        Actions     = $null
    })
}

# Compliance-Richtlinien auf Settings-Catalog-Basis (z. B. Linux)
foreach ($p in (Get-GraphAllSafe -Uri "$graphBeta/deviceManagement/compliancePolicies?`$expand=assignments" -Label 'Compliance-Richtlinien (compliancePolicies)')) {
    $profiles.Add([pscustomobject]@{
        Typ         = 'Compliance (Settings Catalog)'
        Detailtyp   = @($p.technologies) -join ', '
        Name        = $p.name
        Plattform   = Get-NormalizedPlatform $p.platforms
        Id          = $p.id
        Assignments = @($p.assignments)
        ActionsUrl  = "/deviceManagement/compliancePolicies/$($p.id)/scheduledActionsForRule?`$expand=scheduledActionConfigurations"
        Actions     = $null
    })
}

Write-Host "Richtlinien gefunden: $($profiles.Count)" -ForegroundColor Cyan
#endregion

#region Aktionen bei Nichtkonformität per Batch laden
if ($profiles.Count -gt 0) {
    try {
        $actionResp = Invoke-GraphBatch -RelativeUrls @($profiles | ForEach-Object { $_.ActionsUrl }) -Version beta
    }
    catch {
        $msg = "Aktionen bei Nichtkonformität konnten nicht geladen werden: $($_.Exception.Message)"
        Write-Warning $msg; $loadErrors.Add($msg)
        $actionResp = @{}
    }
    foreach ($p in $profiles) {
        $r = $actionResp[$p.ActionsUrl]
        if ($r -and $r.status -eq 200) {
            # Alle Aktionen aller Regeln einer Richtlinie flach sammeln
            $p.Actions = @(foreach ($rule in @($r.body.value)) {
                foreach ($c in @($rule.scheduledActionConfigurations)) { if ($c) { $c } }
            })
        }
        elseif ($r) {
            $msg = "Richtlinie '$($p.Name)': Aktionen nicht lesbar (HTTP $($r.status))"
            Write-Warning $msg; $loadErrors.Add($msg)
        }
    }
}
#endregion

#region Zuweisungen aufbereiten
$rows = foreach ($p in $profiles) {
    $assignments = @($p.Assignments | Where-Object { $_ })

    if ($assignments.Count -eq 0) {
        [pscustomobject]@{
            Richtlinie = $p.Name; Typ = $p.Typ; Plattform = $p.Plattform; Detailtyp = $p.Detailtyp
            Modus = '-'; Bereich = 'Keine'; Gruppe = '<nicht zugewiesen>'
            Filter = $null; FilterModus = $null; NichtKonformNach = $null; Aktionen = $null
            GruppenId = $null; RichtlinienId = $p.Id
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
            Richtlinie       = $p.Name
            Typ              = $p.Typ
            Plattform        = $p.Plattform
            Detailtyp        = $p.Detailtyp
            Modus            = $modus
            Bereich          = $bereich
            Gruppe           = $gruppe
            Filter           = $filter
            FilterModus      = if ($hasFilter) { $t.deviceAndAppManagementAssignmentFilterType } else { $null }
            NichtKonformNach = $null
            Aktionen         = $null
            GruppenId        = $groupId
            RichtlinienId    = $p.Id
        }
    }
}
$rows = @($rows)
#endregion

#region Gruppennamen per Batch auflösen (Zuweisungen + CC-Gruppen der Benachrichtigungen)
$ccGroupIds = @($profiles | ForEach-Object { $_.Actions } | Where-Object { $_ } |
                ForEach-Object { @($_.notificationMessageCCList) } | Where-Object { $_ })
$groupIds   = @(@($rows | Where-Object { $_.GruppenId } | ForEach-Object { $_.GruppenId }) + $ccGroupIds |
                Select-Object -Unique)

$groupNames = @{}
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

#region Aktionen formatieren
$actionText   = @{}
$nonCompliant = @{}
foreach ($p in $profiles) {
    if ($null -eq $p.Actions) {
        $actionText[$p.Id]   = '<nicht lesbar>'
        $nonCompliant[$p.Id] = '<nicht lesbar>'
        continue
    }

    $sorted = @($p.Actions | Sort-Object @{ Expression = { [int]$_.gracePeriodHours } },
                                         @{ Expression = { $i = [array]::IndexOf($actionOrder, "$($_.actionType)"); if ($i -lt 0) { 99 } else { $i } } })

    $labels = foreach ($c in $sorted) {
        $type  = "$($c.actionType)"
        $label = if ($actionLabels.ContainsKey($type)) { $actionLabels[$type] } else { $type }
        $text  = "${label}: $(Format-GracePeriod $c.gracePeriodHours)"

        $details = @()
        if ($c.notificationTemplateId -and $c.notificationTemplateId -ne $emptyGuid) {
            $tpl = if ($templateNames.ContainsKey($c.notificationTemplateId)) { $templateNames[$c.notificationTemplateId] }
                   else { $c.notificationTemplateId }
            $details += "Vorlage: $tpl"
        }
        $cc = @($c.notificationMessageCCList | Where-Object { $_ })
        if ($cc.Count -gt 0) {
            $details += "CC: " + (@($cc | ForEach-Object { if ($groupNames.ContainsKey($_)) { $groupNames[$_] } else { $_ } }) -join ', ')
        }
        if ($details.Count -gt 0) { $text += " ($($details -join '; '))" }
        $text
    }

    $actionText[$p.Id] = if ($labels) { @($labels) -join ' | ' } else { '<keine>' }

    $block = @($sorted | Where-Object { "$($_.actionType)" -eq 'block' }) | Select-Object -First 1
    $nonCompliant[$p.Id] = if ($block) { Format-GracePeriod $block.gracePeriodHours -Short } else { '<keine Aktion>' }
}

foreach ($row in $rows) {
    $row.Aktionen         = $actionText[$row.RichtlinienId]
    $row.NichtKonformNach = $nonCompliant[$row.RichtlinienId]
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
    $result = @($filtered | Sort-Object Richtlinie, RichtlinienId, $modusOrder, Gruppe |
                Select-Object Richtlinie, Typ, Plattform, Detailtyp, Modus, Bereich, Gruppe, Filter, FilterModus,
                              NichtKonformNach, Aktionen, GruppenId, RichtlinienId)
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
            Richtlinie       = $first.Richtlinie
            Typ              = $first.Typ
            Plattform        = $first.Plattform
            Detailtyp        = $first.Detailtyp
            Zuweisungen      = @($items | Where-Object { $_.Bereich -ne 'Keine' }).Count
            Gruppe           = $labels -join ' | '
            NichtKonformNach = $first.NichtKonformNach
            Aktionen         = $first.Aktionen
            GruppenIds       = @($items | Where-Object { $_.GruppenId } | ForEach-Object { $_.GruppenId }) -join ' | '
            RichtlinienId    = $first.RichtlinienId
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
        $result | Format-Table Richtlinie, Plattform, Modus, Gruppe, Filter, FilterModus, NichtKonformNach -AutoSize -Wrap | Out-Host
    }
    else {
        $result | Format-Table Richtlinie, Plattform, Zuweisungen, Gruppe, NichtKonformNach -AutoSize -Wrap | Out-Host
    }

    $exportDir = Split-Path -Path $ExportPath -Parent
    if ($exportDir -and -not (Test-Path -Path $exportDir)) {
        New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
    }
    $encoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
    $result | Export-Csv -Path $ExportPath -Delimiter ';' -NoTypeInformation -Encoding $encoding
    Write-Host "CSV exportiert: $ExportPath ($($result.Count) Zeilen)" -ForegroundColor Green

    if ($GridView) { $result | Out-GridView -Title 'Intune Compliance-Richtlinien – Zuweisungen' }
}

if ($loadErrors.Count -gt 0) {
    Write-Warning "Der Report ist UNVOLLSTÄNDIG – $($loadErrors.Count) Fehler beim Laden:"
    $loadErrors | ForEach-Object { Write-Warning "  $_" }
}

if ($PassThru) { $result }
#endregion
