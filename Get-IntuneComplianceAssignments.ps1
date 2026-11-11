<#
.SYNOPSIS
    Lists all Intune compliance policies including their assignments (group, include/exclude, filter)
    and actions for noncompliance.

.DESCRIPTION
    Covers the following policy types via Microsoft Graph (beta):
      - Compliance policies (classic)                 /deviceManagement/deviceCompliancePolicies
      - Compliance policies (Settings Catalog based)  /deviceManagement/compliancePolicies   (e.g. Linux)

    For every policy the actions for noncompliance are returned as well (e.g. mark device
    noncompliant, send email, remote lock, add to retire list) incl. timing, notification template
    and CC groups. In addition, the script shows the tenant-wide compliance settings
    ("Mark devices with no compliance policy assigned as ..." and the compliance status validity period).

    Errors while loading a single policy type do not abort the run; they are reported as warnings
    and summarized at the end. Group names and actions are resolved via JSON batching
    ($batch, 20 requests per call, retry on HTTP 429/503/504).

    Required permissions (delegated or application):
      DeviceManagementConfiguration.Read.All (or ReadWrite.All)
      Group.Read.All (or Group.ReadWrite.All / Directory.Read.All)
    Optional (only for notification template names and tenant-wide settings):
      DeviceManagementServiceConfig.Read.All – if missing, template IDs are shown instead of names
      and the tenant-wide settings are skipped with a warning. The report is NOT marked incomplete.

    Sign-in:
      Without an existing connection the script signs in interactively (delegated).
      For app-only (e.g. automation) connect yourself beforehand, e.g.:
        Connect-MgGraph -TenantId <tenant> -ClientId <appId> -CertificateThumbprint <thumbprint>
      The script then verifies the application permissions with test requests.

    Output format (default): ONE row per policy. All assignments are listed in the column
    "AssignedGroups", separated by " | ", includes before excludes, filters inline, e.g.:
      GRP-Win-Pilot | GRP-Win-Prod [Filter: Laptops only (include)] | Exclude: GRP-Win-Kiosk
    The column "Actions" contains all actions for noncompliance in chronological order, e.g.:
      Mark device noncompliant: immediately | Send email to end user: after 1 day (Template: Notice) | ...
    The column "NoncompliantAfter" shows separately when a device is marked as noncompliant.
    The console table omits "Actions" for space reasons – it is included in CSV, GridView and -PassThru.
    With -OneRowPerAssignment you get one row per assignment instead, with separate columns
    for Mode, Filter and FilterMode.

.PARAMETER TenantId
    Optional: tenant ID or any verified domain of the tenant. Forces the connection to exactly this tenant
    (important for multi-tenant admins / guest accounts).

.PARAMETER GroupName
    Optional: answers the question "Which compliance policies are assigned to this group?".
    Returns policies with at least one direct assignment to a group whose name matches the pattern
    (wildcards supported; include or exclude). For these policies ALL assignments are shown
    (other groups, excludes and tenant-wide assignments as well).
    Nested group membership is not resolved. Policies without assignments never appear
    when -GroupName is used.

.PARAMETER IncludeTenantWide
    Only with -GroupName: additionally return all policies assigned to "All devices" or "All users",
    even without a match on the group.

.PARAMETER IncludeUnassigned
    Also list policies without assignments. Has no effect with -GroupName.

.PARAMETER OneRowPerAssignment
    One row per assignment instead of one row per policy (with columns Mode, Filter, FilterMode) –
    useful for pivot tables in Excel.

.PARAMETER ExportPath
    Path of the CSV file (UTF-8 with BOM). The target folder is created if missing.
    Default: %TEMP%\Intune-ComplianceAssignments_<timestamp>.csv

.PARAMETER Delimiter
    CSV delimiter. Default ';' (opens directly in Excel with German/European regional settings).
    Use ',' for English regional settings.

.PARAMETER GridView
    Additionally show the result in Out-GridView.

.PARAMETER PassThru
    Return the result objects to the pipeline (instead of table/CSV only).

.EXAMPLE
    .\Get-IntuneComplianceAssignments.ps1 -TenantId contoso.onmicrosoft.com
.EXAMPLE
    .\Get-IntuneComplianceAssignments.ps1 -GroupName "GRP-Intune-Win-*" -IncludeTenantWide -GridView
.EXAMPLE
    .\Get-IntuneComplianceAssignments.ps1 -IncludeUnassigned -ExportPath C:\Reports\compliance.csv
.EXAMPLE
    $r = .\Get-IntuneComplianceAssignments.ps1 -PassThru; $r | Where-Object NoncompliantAfter -ne 'immediately'
.EXAMPLE
    .\Get-IntuneComplianceAssignments.ps1 -OneRowPerAssignment -Delimiter ','

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
    [string]$GroupName,
    [switch]$IncludeTenantWide,
    [switch]$IncludeUnassigned,
    [string]$ExportPath = (Join-Path $env:TEMP "Intune-ComplianceAssignments_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"),
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
    # Application permissions are not reliably visible in the context -> functional test via probe requests
    $probes = [ordered]@{
        'DeviceManagementConfiguration.Read.All' = "$graphBeta/deviceManagement/deviceCompliancePolicies?`$top=1&`$select=id"
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

function Get-NormalizedPlatform {
    param([string]$Value)
    if (-not $Value) { return 'Unknown' }
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
    # Default: "after 3 days" (for the action list); -Short: "3 days" (for the column NoncompliantAfter)
    param($Hours, [switch]$Short)
    $h = [int]$Hours
    if ($h -le 0) { return 'immediately' }
    if ($h % 24 -ne 0) {
        $text = if ($h -eq 1) { '1 hour' } else { "$h hours" }
    }
    else {
        $d    = $h / 24
        $text = if ($d -eq 1) { '1 day' } else { "$d days" }
    }
    if ($Short) { return $text }
    return "after $text"
}

# Display names of the actions (as in the Intune admin center)
$actionLabels = @{
    'block'                        = 'Mark device noncompliant'
    'notification'                 = 'Send email to end user'
    'pushNotification'             = 'Send push notification to end user'
    'remoteLock'                   = 'Remotely lock the noncompliant device'
    'retire'                       = 'Add device to retire list'
    'wipe'                         = 'Wipe device'
    'removeResourceAccessProfiles' = 'Remove resource access profiles'
    'noAction'                     = 'No action'
}

# Order for actions with the same timing (hashtable keys are case-insensitive)
$actionRank = @{}
$rank = 0
foreach ($a in 'block', 'notification', 'pushNotification', 'remoteLock', 'retire', 'wipe', 'removeResourceAccessProfiles', 'noAction') {
    $actionRank[$a] = $rank++
}
#endregion

#region Load tenant-wide settings, filters and notification templates
# Not every request shape returns the "settings" property (e.g. $select is ignored in some tenants),
# so several variants are tried until one actually contains it
$s = $null
$settingsError = $null
foreach ($uri in "$graphBeta/deviceManagement?`$select=settings",
                 "$graphBeta/deviceManagement",
                 'https://graph.microsoft.com/v1.0/deviceManagement') {
    try {
        $dm = Invoke-MgGraphRequest -Method GET -Uri $uri
        if ($dm.settings) { $s = $dm.settings; break }
    }
    catch {
        $settingsError = $_.Exception.Message
    }
}

if ($s) {
    # Missing single values are reported as "unknown" instead of being silently interpreted as "Compliant"
    $noPolicyState = if     ($null -eq $s.secureByDefault) { 'unknown (not returned)' }
                     elseif ($s.secureByDefault)           { 'Not compliant' }
                     else                                  { 'Compliant' }
    $validity      = if ($null -eq $s.deviceComplianceCheckinThresholdDays) { 'unknown (not returned)' }
                     else { "$($s.deviceComplianceCheckinThresholdDays) days" }
    Write-Host ""
    Write-Host "Tenant-wide compliance settings" -ForegroundColor Cyan
    Write-Host "  Mark devices with no compliance policy assigned as: $noPolicyState"
    Write-Host "  Compliance status validity period:                  $validity"
    Write-Host ""
}
else {
    $detail = if ($settingsError) { " ($settingsError)" } else { ' (property not returned by Graph)' }
    Write-Warning "Tenant-wide compliance settings not readable$detail. May require DeviceManagementServiceConfig.Read.All."
}

$filterNames = @{}
foreach ($f in (Get-GraphAllSafe -Uri "$graphBeta/deviceManagement/assignmentFilters" -Label 'Assignment filters')) {
    $filterNames[$f.id] = $f.displayName
}

# Template names are optional (separate permission) – errors do not make the report incomplete
$templateNames = @{}
try {
    foreach ($t in (Get-GraphAll -Uri "$graphBeta/deviceManagement/notificationMessageTemplates?`$select=id,displayName")) {
        $templateNames[$t.id] = $t.displayName
    }
}
catch {
    Write-Warning "Notification template names not readable – template IDs are shown instead. May require DeviceManagementServiceConfig.Read.All."
}
#endregion

#region Load policies
$profiles = [System.Collections.Generic.List[object]]::new()

# Compliance policies (classic)
foreach ($p in (Get-GraphAllSafe -Uri "$graphBeta/deviceManagement/deviceCompliancePolicies?`$expand=assignments" -Label 'Compliance policies (deviceCompliancePolicies)')) {
    $odata = $p.'@odata.type' -replace '^#microsoft\.graph\.', ''
    $profiles.Add([pscustomobject]@{
        Type        = 'Compliance'
        DetailType  = $odata
        Name        = $p.displayName
        Platform    = Get-NormalizedPlatform $odata
        Id          = $p.id
        Assignments = @($p.assignments)
        BasePath    = "/deviceManagement/deviceCompliancePolicies/$($p.id)"
        Actions     = $null
    })
}

# Compliance policies based on the Settings Catalog (e.g. Linux)
foreach ($p in (Get-GraphAllSafe -Uri "$graphBeta/deviceManagement/compliancePolicies?`$expand=assignments" -Label 'Compliance policies (compliancePolicies)')) {
    $profiles.Add([pscustomobject]@{
        Type        = 'Compliance (Settings Catalog)'
        DetailType  = @($p.technologies) -join ', '
        Name        = $p.name
        Platform    = Get-NormalizedPlatform $p.platforms
        Id          = $p.id
        Assignments = @($p.assignments)
        BasePath    = "/deviceManagement/compliancePolicies/$($p.id)"
        Actions     = $null
    })
}

Write-Host "Policies found: $($profiles.Count)" -ForegroundColor Cyan
#endregion

#region Load actions for noncompliance via batch
# Graph does not accept every query shape for this relationship in every tenant (observed: HTTP 400
# for the expanded list request). Three variants are tried in order, each only for the policies
# that are still open:
#   1. <policy>/scheduledActionsForRule?$expand=scheduledActionConfigurations
#   2. <policy>?$expand=scheduledActionsForRule($expand=scheduledActionConfigurations)
#   3. <policy>/scheduledActionsForRule, then <rule>/scheduledActionConfigurations per rule
function Get-BatchError {
    param($Response)
    if (-not $Response) { return 'no response' }
    if ([int]$Response.status -eq 0) { return 'batch request failed' }
    $m = $Response.body.error.message
    if ($m) { return "HTTP $($Response.status): $m" }
    return "HTTP $($Response.status)"
}

$pending   = @($profiles)
$lastError = @{}

$variants = @(
    @{ Url   = { param($p) "$($p.BasePath)/scheduledActionsForRule?`$expand=scheduledActionConfigurations" }
       Rules = { param($b) @($b.value) } },
    @{ Url   = { param($p) "$($p.BasePath)?`$expand=scheduledActionsForRule(`$expand=scheduledActionConfigurations)" }
       Rules = { param($b) @($b.scheduledActionsForRule) } }
)

foreach ($v in $variants) {
    if ($pending.Count -eq 0) { break }
    $urls = @{}
    foreach ($p in $pending) { $urls[$p.Id] = & $v.Url $p }
    try {
        $resp = Invoke-GraphBatch -RelativeUrls @($urls.Values) -Version beta
    }
    catch {
        $err = $_.Exception.Message
        foreach ($p in $pending) { $lastError[$p.Id] = $err }
        continue
    }
    $still = foreach ($p in $pending) {
        $r = $resp[$urls[$p.Id]]
        if ($r -and $r.status -eq 200) {
            # Collect all actions of all rules of a policy in one flat list
            $p.Actions = @(foreach ($rule in (& $v.Rules $r.body)) {
                foreach ($c in @($rule.scheduledActionConfigurations)) { if ($c) { $c } }
            })
        }
        else {
            $lastError[$p.Id] = Get-BatchError $r
            $p
        }
    }
    $pending = @($still)
}

# Variant 3: rules first, then the configurations of every rule
if ($pending.Count -gt 0) {
    $ruleUrls = @{}
    foreach ($p in $pending) { $ruleUrls[$p.Id] = "$($p.BasePath)/scheduledActionsForRule" }
    try   { $ruleResp = Invoke-GraphBatch -RelativeUrls @($ruleUrls.Values) -Version beta }
    catch { $ruleResp = @{}; $err = $_.Exception.Message; foreach ($p in $pending) { $lastError[$p.Id] = $err } }

    $cfgUrlsByPolicy = @{}
    foreach ($p in $pending) {
        $r = $ruleResp[$ruleUrls[$p.Id]]
        if ($r -and $r.status -eq 200) {
            $cfgUrlsByPolicy[$p.Id] = @(@($r.body.value) | Where-Object { $_ } |
                ForEach-Object { "$($p.BasePath)/scheduledActionsForRule/$($_.id)/scheduledActionConfigurations" })
        }
        elseif ($r) {
            $lastError[$p.Id] = Get-BatchError $r
        }
    }

    $allCfgUrls = @($cfgUrlsByPolicy.Values | ForEach-Object { $_ })
    $cfgResp = @{}
    if ($allCfgUrls.Count -gt 0) {
        try   { $cfgResp = Invoke-GraphBatch -RelativeUrls $allCfgUrls -Version beta }
        catch { $err = $_.Exception.Message; foreach ($id in $cfgUrlsByPolicy.Keys) { $lastError[$id] = $err } }
    }

    $still = foreach ($p in $pending) {
        if (-not $cfgUrlsByPolicy.ContainsKey($p.Id)) { $p; continue }
        $ok   = $true
        $acts = foreach ($u in $cfgUrlsByPolicy[$p.Id]) {
            $r = $cfgResp[$u]
            if ($r -and $r.status -eq 200) { @($r.body.value) }
            else { $ok = $false; $lastError[$p.Id] = Get-BatchError $r }
        }
        if ($ok) { $p.Actions = @($acts | Where-Object { $_ }) } else { $p }
    }
    $pending = @($still)
}

# Remaining failures: one warning per distinct error instead of one per policy
if ($pending.Count -gt 0) {
    foreach ($g in ($pending | Group-Object { $lastError[$_.Id] })) {
        $msg = "Actions for noncompliance not readable for $($g.Count) policy(ies) [$($g.Name)]: $(@($g.Group | ForEach-Object { $_.Name }) -join ', ')"
        Write-Warning $msg; $loadErrors.Add($msg)
    }
}
#endregion

#region Build assignment rows
$rows = foreach ($p in $profiles) {
    $assignments = @($p.Assignments | Where-Object { $_ })

    if ($assignments.Count -eq 0) {
        [pscustomobject]@{
            Policy = $p.Name; Type = $p.Type; Platform = $p.Platform; DetailType = $p.DetailType
            Mode = '-'; Scope = 'None'; Group = '<not assigned>'
            Filter = $null; FilterMode = $null; NoncompliantAfter = $null; Actions = $null
            GroupId = $null; PolicyId = $p.Id
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
            Policy            = $p.Name
            Type              = $p.Type
            Platform          = $p.Platform
            DetailType        = $p.DetailType
            Mode              = $mode
            Scope             = $scope
            Group             = $group
            Filter            = $filter
            FilterMode        = if ($hasFilter) { $t.deviceAndAppManagementAssignmentFilterType } else { $null }
            NoncompliantAfter = $null
            Actions           = $null
            GroupId           = $groupId
            PolicyId          = $p.Id
        }
    }
}
$rows = @($rows)
#endregion

#region Resolve group names via batch (assignments + CC groups of notifications)
$ccGroupIds = @($profiles | ForEach-Object { $_.Actions } | Where-Object { $_ } |
                ForEach-Object { @($_.notificationMessageCCList) } | Where-Object { $_ })
$groupIds   = @(@($rows | Where-Object { $_.GroupId } | ForEach-Object { $_.GroupId }) + $ccGroupIds |
                Select-Object -Unique)

$groupNames = @{}
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

#region Format actions
$actionText   = @{}
$nonCompliant = @{}
foreach ($p in $profiles) {
    if ($null -eq $p.Actions) {
        $actionText[$p.Id]   = '<not readable>'
        $nonCompliant[$p.Id] = '<not readable>'
        continue
    }

    $sorted = @($p.Actions | Sort-Object @{ Expression = { [int]$_.gracePeriodHours } },
                                         @{ Expression = { $k = "$($_.actionType)"; if ($actionRank.ContainsKey($k)) { $actionRank[$k] } else { 99 } } })

    $labels = foreach ($c in $sorted) {
        $type  = "$($c.actionType)"
        $label = if ($actionLabels.ContainsKey($type)) { $actionLabels[$type] } else { $type }
        # "No action" has no meaningful timing
        $text  = if ($type -eq 'noAction') { $label } else { "${label}: $(Format-GracePeriod $c.gracePeriodHours)" }

        $details = @()
        if ($c.notificationTemplateId -and $c.notificationTemplateId -ne $emptyGuid) {
            $tpl = if ($templateNames.ContainsKey($c.notificationTemplateId)) { $templateNames[$c.notificationTemplateId] }
                   else { $c.notificationTemplateId }
            $details += "Template: $tpl"
        }
        $cc = @($c.notificationMessageCCList | Where-Object { $_ })
        if ($cc.Count -gt 0) {
            $details += "CC: " + (@($cc | ForEach-Object { if ($groupNames.ContainsKey($_)) { $groupNames[$_] } else { $_ } }) -join ', ')
        }
        if ($details.Count -gt 0) { $text += " ($($details -join '; '))" }
        $text
    }

    $actionText[$p.Id] = if ($labels) { @($labels) -join ' | ' } else { '<none>' }

    $block = @($sorted | Where-Object { "$($_.actionType)" -eq 'block' }) | Select-Object -First 1
    $nonCompliant[$p.Id] = if ($block) { Format-GracePeriod $block.gracePeriodHours -Short } else { '<no action>' }
}

foreach ($row in $rows) {
    $row.Actions           = $actionText[$row.PolicyId]
    $row.NoncompliantAfter = $nonCompliant[$row.PolicyId]
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
    $result = @($filtered | Sort-Object Policy, PolicyId, $modeOrder, Group |
                Select-Object Policy, Type, Platform, DetailType, Mode, Scope, Group, Filter, FilterMode,
                              NoncompliantAfter, Actions, GroupId, PolicyId)
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
            Policy            = $first.Policy
            Type              = $first.Type
            Platform          = $first.Platform
            DetailType        = $first.DetailType
            AssignmentCount   = @($items | Where-Object { $_.Scope -ne 'None' }).Count
            AssignedGroups    = $labels -join ' | '
            NoncompliantAfter = $first.NoncompliantAfter
            Actions           = $first.Actions
            GroupIds          = @($items | Where-Object { $_.GroupId } | ForEach-Object { $_.GroupId }) -join ' | '
            PolicyId          = $first.PolicyId
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
        $result | Format-Table Policy, Platform, Mode, Group, Filter, FilterMode, NoncompliantAfter -AutoSize -Wrap | Out-Host
    }
    else {
        $result | Format-Table Policy, Platform, AssignmentCount, AssignedGroups, NoncompliantAfter -AutoSize -Wrap | Out-Host
    }

    $exportDir = Split-Path -Path $ExportPath -Parent
    if ($exportDir -and -not (Test-Path -Path $exportDir)) {
        New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
    }
    $encoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
    $result | Export-Csv -Path $ExportPath -Delimiter $Delimiter -NoTypeInformation -Encoding $encoding
    Write-Host "CSV exported: $ExportPath ($($result.Count) rows)" -ForegroundColor Green

    if ($GridView) { $result | Out-GridView -Title 'Intune compliance policies – assignments' }
}

if ($loadErrors.Count -gt 0) {
    Write-Warning "The report is INCOMPLETE – $($loadErrors.Count) error(s) while loading:"
    $loadErrors | ForEach-Object { Write-Warning "  $_" }
}

if ($PassThru) { $result }
#endregion
