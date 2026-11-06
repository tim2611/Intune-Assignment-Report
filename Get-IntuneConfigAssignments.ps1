<#
.SYNOPSIS
    Lists all Intune configuration profiles including their assignments (group, include/exclude, filter).

.DESCRIPTION
    Covers the following profile types via Microsoft Graph (beta):
      - Templates (classic)                              /deviceManagement/deviceConfigurations
      - Settings Catalog / Endpoint Security / Baselines /deviceManagement/configurationPolicies
      - Administrative Templates (ADMX)                  /deviceManagement/groupPolicyConfigurations
      - Endpoint Security (legacy intents)               /deviceManagement/intents

    Errors while loading a single profile type do not abort the run; they are reported as warnings
    and summarized at the end. Group names and intent assignments are resolved via JSON batching
    ($batch, 20 requests per call, retry on HTTP 429/503/504).

    Required permissions (delegated or application):
      DeviceManagementConfiguration.Read.All (or ReadWrite.All)
      Group.Read.All (or Group.ReadWrite.All / Directory.Read.All)

    Sign-in:
      Without an existing connection the script signs in interactively (delegated).
      For app-only (e.g. automation) connect yourself beforehand, e.g.:
        Connect-MgGraph -TenantId <tenant> -ClientId <appId> -CertificateThumbprint <thumbprint>
      The script then verifies the application permissions with test requests.

    Output format (default): ONE row per policy. All assignments are listed in the column
    "AssignedGroups", separated by " | ", includes before excludes, filters inline, e.g.:
      GRP-Win-Pilot | GRP-Win-Prod [Filter: Laptops only (include)] | Exclude: GRP-Win-Kiosk
    The column "AssignmentCount" counts all assignments of the policy (include, exclude,
    tenant-wide, unknown target type).
    With -OneRowPerAssignment you get one row per assignment instead, with separate columns
    for Mode, Filter and FilterMode.

.PARAMETER TenantId
    Optional: tenant ID or primary domain. Forces the connection to exactly this tenant
    (important for multi-tenant admins / guest accounts).

.PARAMETER GroupName
    Optional: answers the question "Which policies are assigned to this group?".
    Returns policies with at least one direct assignment to a group whose name matches the pattern
    (wildcards supported; include or exclude). For these policies ALL assignments are shown
    (other groups, excludes and tenant-wide assignments as well).
    Nested group membership is not resolved. Policies without assignments never appear
    when -GroupName is used.

.PARAMETER IncludeTenantWide
    Only with -GroupName: additionally return all policies assigned to "All devices" or "All users",
    even without a match on the group. The report then contains everything members of the group can
    receive via direct AND tenant-wide assignments (mind excludes and filters).

.PARAMETER IncludeUnassigned
    Also list policies without assignments. Has no effect with -GroupName.

.PARAMETER OneRowPerAssignment
    One row per assignment instead of one row per policy (with columns Mode, Filter, FilterMode) –
    useful for pivot tables in Excel.

.PARAMETER ExportPath
    Path of the CSV file (UTF-8 with BOM). The target folder is created if missing.
    Default: %TEMP%\Intune-ConfigAssignments_<timestamp>.csv

.PARAMETER Delimiter
    CSV delimiter. Default ';' (opens directly in Excel with German/European regional settings).
    Use ',' for English regional settings.

.PARAMETER GridView
    Additionally show the result in Out-GridView.

.PARAMETER PassThru
    Return the result objects to the pipeline (instead of table/CSV only).

.EXAMPLE
    .\Get-IntuneConfigAssignments.ps1 -TenantId contoso.onmicrosoft.com
.EXAMPLE
    .\Get-IntuneConfigAssignments.ps1 -GroupName "GRP-Intune-Win-*" -IncludeTenantWide -GridView
.EXAMPLE
    .\Get-IntuneConfigAssignments.ps1 -IncludeUnassigned -ExportPath C:\Reports\assignments.csv
.EXAMPLE
    $r = .\Get-IntuneConfigAssignments.ps1 -PassThru; $r | Where-Object AssignmentCount -gt 1
.EXAMPLE
    .\Get-IntuneConfigAssignments.ps1 -OneRowPerAssignment -Delimiter ','

.NOTES
    Author:   Timo Schlappinger
    Version:  1.1.0
    Project:  https://github.com/tim2611/Intune-Assignment-Report
    Requires: Module Microsoft.Graph.Authentication, PowerShell 5.1 or 7.x
#>
[CmdletBinding()]
[OutputType([pscustomobject])]
param(
    [string]$TenantId,
    [string]$GroupName,
    [switch]$IncludeTenantWide,
    [switch]$IncludeUnassigned,
    [string]$ExportPath = (Join-Path $env:TEMP "Intune-ConfigAssignments_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"),
    [ValidateSet(';', ',', "`t")][string]$Delimiter = ';',
    [switch]$OneRowPerAssignment,
    [switch]$GridView,
    [switch]$PassThru
)

$ErrorActionPreference = 'Stop'
$graphBeta  = 'https://graph.microsoft.com/beta'
$emptyGuid  = '00000000-0000-0000-0000-000000000000'
$loadErrors = [System.Collections.Generic.List[string]]::new()

#region Connection & permissions
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    throw "Module missing: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
}
Import-Module Microsoft.Graph.Authentication

# Required permission -> accepted (equal or higher) permissions
$scopeMap = [ordered]@{
    'DeviceManagementConfiguration.Read.All' = @('DeviceManagementConfiguration.Read.All', 'DeviceManagementConfiguration.ReadWrite.All')
    'Group.Read.All'                         = @('Group.Read.All', 'Group.ReadWrite.All', 'Directory.Read.All', 'Directory.ReadWrite.All')
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

$ctx         = Get-MgContext
$isAppOnly   = $ctx -and ("$($ctx.AuthType)" -eq 'AppOnly')
$wrongTenant = $ctx -and $TenantId -and ($ctx.TenantId -ne $TenantId)

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
    # Application permissions are not reliably visible in the context -> functional test via probe requests
    $probes = [ordered]@{
        'DeviceManagementConfiguration.Read.All' = "$graphBeta/deviceManagement/deviceConfigurations?`$top=1&`$select=id"
        'Group.Read.All'                         = 'https://graph.microsoft.com/v1.0/groups?$top=1&$select=id'
    }
    $failed = foreach ($perm in $probes.Keys) {
        try   { Invoke-MgGraphRequest -Method GET -Uri $probes[$perm] | Out-Null }
        catch { $perm }
    }
    if ($failed) {
        throw "App registration is (probably) missing the application permission: $(@($failed) -join ', ') (incl. admin consent)"
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

function Get-GraphAllSafe {
    param([Parameter(Mandatory)][string]$Uri, [Parameter(Mandatory)][string]$Label)
    try {
        Get-GraphAll -Uri $Uri
    }
    catch {
        $msg = "$Label could not be loaded: $($_.Exception.Message)"
        Write-Warning $msg
        $script:loadErrors.Add($msg)
    }
}

function Invoke-GraphBatch {
    <#
      Runs GET requests via JSON batching (max. 20 per call).
      Single 429/503/504 responses inside a batch are retried up to 3 times honoring Retry-After,
      as are 5xx/transport errors of the batch POST itself. Status 0 = batch POST never succeeded.
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

        # Requests still open after all attempts: keep their last response (status)
        foreach ($url in $pending) {
            $results[$url] = if ($lastResp.ContainsKey($url)) { $lastResp[$url] } else { @{ status = 0; body = $null } }
        }
    }
    return $results
}

function Get-NormalizedPlatform {
    param([string]$Value)
    if (-not $Value) { return 'Unknown' }
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

#region Load assignment filters
$filterNames = @{}
foreach ($f in (Get-GraphAllSafe -Uri "$graphBeta/deviceManagement/assignmentFilters" -Label 'Assignment filters')) {
    $filterNames[$f.id] = $f.displayName
}
#endregion

#region Load profiles
$profiles = [System.Collections.Generic.List[object]]::new()

# Templates (classic)
foreach ($p in (Get-GraphAllSafe -Uri "$graphBeta/deviceManagement/deviceConfigurations?`$expand=assignments" -Label 'Templates (deviceConfigurations)')) {
    $odata = $p.'@odata.type' -replace '^#microsoft\.graph\.', ''
    $profiles.Add([pscustomobject]@{
        Type        = 'Template'
        DetailType  = $odata
        Name        = $p.displayName
        Platform    = Get-NormalizedPlatform $odata
        Id          = $p.id
        Assignments = @($p.assignments)
    })
}

# Settings Catalog + new Endpoint Security policies + new Security Baselines
foreach ($p in (Get-GraphAllSafe -Uri "$graphBeta/deviceManagement/configurationPolicies?`$expand=assignments" -Label 'Settings Catalog (configurationPolicies)')) {
    $family = $p.templateReference.templateFamily
    $isTemplate = $family -and $family -ne 'none'
    $profiles.Add([pscustomobject]@{
        Type        = if ($isTemplate) { 'Endpoint Security / Baseline' } else { 'Settings Catalog' }
        DetailType  = if ($isTemplate) { $family } else { @($p.technologies) -join ', ' }
        Name        = $p.name
        Platform    = Get-NormalizedPlatform $p.platforms
        Id          = $p.id
        Assignments = @($p.assignments)
    })
}

# Administrative Templates (ADMX)
foreach ($p in (Get-GraphAllSafe -Uri "$graphBeta/deviceManagement/groupPolicyConfigurations?`$expand=assignments" -Label 'Administrative Templates (groupPolicyConfigurations)')) {
    $profiles.Add([pscustomobject]@{
        Type        = 'Administrative Template'
        DetailType  = 'groupPolicyConfiguration'
        Name        = $p.displayName
        Platform    = 'Windows'
        Id          = $p.id
        Assignments = @($p.assignments)
    })
}

# Endpoint Security legacy intents (no $expand possible -> assignments via batch)
$intents = @(Get-GraphAllSafe -Uri "$graphBeta/deviceManagement/intents" -Label 'Endpoint Security legacy intents')
if ($intents.Count -gt 0) {
    $intentUrls = @($intents | ForEach-Object { "/deviceManagement/intents/$($_.id)/assignments" })
    try {
        $intentAssignments = Invoke-GraphBatch -RelativeUrls $intentUrls -Version beta
    }
    catch {
        $msg = "Assignments of legacy intents could not be loaded: $($_.Exception.Message)"
        Write-Warning $msg; $loadErrors.Add($msg)
        $intentAssignments = @{}
    }
    foreach ($p in $intents) {
        $r = $intentAssignments["/deviceManagement/intents/$($p.id)/assignments"]
        if ($r -and $r.status -ne 200) {
            $msg = "Intent '$($p.displayName)': assignments not readable (HTTP $($r.status))"
            Write-Warning $msg; $loadErrors.Add($msg)
        }
        $profiles.Add([pscustomobject]@{
            Type        = 'Endpoint Security (legacy intent)'
            DetailType  = $p.templateId
            Name        = $p.displayName
            Platform    = 'Windows'
            Id          = $p.id
            Assignments = if ($r -and $r.status -eq 200) { @($r.body.value) } else { @() }
        })
    }
}

Write-Host "Profiles found: $($profiles.Count)" -ForegroundColor Cyan
#endregion

#region Build assignment rows
$rows = foreach ($p in $profiles) {
    $assignments = @($p.Assignments | Where-Object { $_ })

    if ($assignments.Count -eq 0) {
        [pscustomobject]@{
            Policy = $p.Name; Type = $p.Type; Platform = $p.Platform; DetailType = $p.DetailType
            Mode = '-'; Scope = 'None'; Group = '<not assigned>'
            Filter = $null; FilterMode = $null; GroupId = $null; PolicyId = $p.Id
        }
        continue
    }

    foreach ($a in $assignments) {
        $t = $a.target
        $groupId = $null; $group = $null; $scope = 'Group'

        switch ($t.'@odata.type') {
            '#microsoft.graph.groupAssignmentTarget'            { $mode = 'Include'; $groupId = $t.groupId }
            '#microsoft.graph.exclusionGroupAssignmentTarget'   { $mode = 'Exclude'; $groupId = $t.groupId }
            '#microsoft.graph.allDevicesAssignmentTarget'       { $mode = 'Include'; $group = 'All devices'; $scope = 'TenantWide' }
            '#microsoft.graph.allLicensedUsersAssignmentTarget' { $mode = 'Include'; $group = 'All users';   $scope = 'TenantWide' }
            default                                             { $mode = 'Unknown'; $group = $t.'@odata.type'; $scope = 'Unknown' }
        }

        $filterId  = $t.deviceAndAppManagementAssignmentFilterId
        $hasFilter = $filterId -and $filterId -ne $emptyGuid
        $filter    = if (-not $hasFilter) { $null }
                     elseif ($filterNames.ContainsKey($filterId)) { $filterNames[$filterId] }
                     else { "<filter deleted/unknown: $filterId>" }

        [pscustomobject]@{
            Policy     = $p.Name
            Type       = $p.Type
            Platform   = $p.Platform
            DetailType = $p.DetailType
            Mode       = $mode
            Scope      = $scope
            Group      = $group
            Filter     = $filter
            FilterMode = if ($hasFilter) { $t.deviceAndAppManagementAssignmentFilterType } else { $null }
            GroupId    = $groupId
            PolicyId   = $p.Id
        }
    }
}
$rows = @($rows)
#endregion

#region Resolve group names via batch
$groupIds = @($rows | Where-Object { $_.GroupId } | Select-Object -ExpandProperty GroupId -Unique)
if ($groupIds.Count -gt 0) {
    $groupUrls = @($groupIds | ForEach-Object { "/groups/$_`?`$select=id,displayName" })
    try {
        $groupResp = Invoke-GraphBatch -RelativeUrls $groupUrls -Version v1.0
    }
    catch {
        $msg = "Group names could not be resolved: $($_.Exception.Message)"
        Write-Warning $msg; $loadErrors.Add($msg)
        $groupResp = @{}
    }
    $groupNames = @{}
    foreach ($id in $groupIds) {
        $r = $groupResp["/groups/$id`?`$select=id,displayName"]
        $groupNames[$id] = if     (-not $r)            { "<not resolved: $id>" }
                           elseif ($r.status -eq 200)  { $r.body.displayName }
                           elseif ($r.status -eq 404)  { "<group deleted: $id>" }
                           else                        { "<error HTTP $($r.status): $id>" }
    }
    foreach ($row in $rows) {
        if ($row.GroupId) { $row.Group = $groupNames[$row.GroupId] }
    }
}
#endregion

#region Filter
$filtered = $rows
if ($GroupName) {
    # Determine policies with at least one matching assignment ...
    $matchingIds = @($rows | Where-Object {
        ($_.Scope -eq 'Group' -and $_.Group -like $GroupName) -or
        ($IncludeTenantWide -and $_.Scope -eq 'TenantWide')
    } | Select-Object -ExpandProperty PolicyId -Unique)
    # ... and keep ALL assignments of these policies (full picture incl. excludes)
    $filtered = $rows | Where-Object { $_.PolicyId -in $matchingIds }
}
elseif (-not $IncludeUnassigned) {
    $filtered = $rows | Where-Object { $_.Scope -ne 'None' }
}
$filtered = @($filtered)

$modeOrder = @{ Expression = { switch ($_.Mode) { 'Include' { 0 } 'Exclude' { 1 } default { 2 } } } }
#endregion

#region Build result
if ($OneRowPerAssignment) {
    # One row per assignment (raw data, e.g. for pivot tables)
    $result = @($filtered | Sort-Object Policy, PolicyId, $modeOrder, Group)
}
else {
    # One row per policy, all assignments combined in the column "AssignedGroups"
    $result = @($filtered | Group-Object PolicyId | ForEach-Object {
        $items = @($_.Group | Sort-Object $modeOrder, Group)
        $first = $items[0]

        $labels = foreach ($i in $items) {
            $label = if ($i.Mode -eq 'Exclude') { "Exclude: $($i.Group)" } else { $i.Group }
            if ($i.Filter) { $label += " [Filter: $($i.Filter) ($($i.FilterMode))]" }
            $label
        }

        [pscustomobject]@{
            Policy          = $first.Policy
            Type            = $first.Type
            Platform        = $first.Platform
            DetailType      = $first.DetailType
            AssignmentCount = @($items | Where-Object { $_.Scope -ne 'None' }).Count
            AssignedGroups  = $labels -join ' | '
            GroupIds        = @($items | Where-Object { $_.GroupId } | ForEach-Object { $_.GroupId }) -join ' | '
            PolicyId        = $first.PolicyId
        }
    } | Sort-Object Policy, PolicyId)
}
#endregion

#region Output
if ($result.Count -eq 0) {
    Write-Warning "No assignments found matching the criteria."
}
else {
    if ($OneRowPerAssignment) {
        $result | Format-Table Policy, Type, Platform, Mode, Group, Filter, FilterMode -AutoSize -Wrap | Out-Host
    }
    else {
        $result | Format-Table Policy, Type, Platform, AssignmentCount, AssignedGroups -AutoSize -Wrap | Out-Host
    }

    $exportDir = Split-Path -Path $ExportPath -Parent
    if ($exportDir -and -not (Test-Path -Path $exportDir)) {
        New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
    }
    $encoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
    $result | Export-Csv -Path $ExportPath -Delimiter $Delimiter -NoTypeInformation -Encoding $encoding
    Write-Host "CSV exported: $ExportPath ($($result.Count) rows)" -ForegroundColor Green

    if ($GridView) { $result | Out-GridView -Title 'Intune configuration profiles – assignments' }
}

if ($loadErrors.Count -gt 0) {
    Write-Warning "The report is INCOMPLETE – $($loadErrors.Count) error(s) while loading:"
    $loadErrors | ForEach-Object { Write-Warning "  $_" }
}

if ($PassThru) { $result }
#endregion
