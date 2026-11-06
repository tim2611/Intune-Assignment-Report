# Intune Assignment Report

PowerShell reports for Microsoft Intune that show **which policies are assigned to which groups** – including include/exclude assignments, tenant-wide assignments and assignment filters – via Microsoft Graph.

| Script | Covers | Status |
|---|---|---|
| [`Get-IntuneConfigAssignments.ps1`](Get-IntuneConfigAssignments.ps1) | Configuration profiles | ✅ Available |
| [`Get-IntuneComplianceAssignments.ps1`](#get-intunecomplianceassignmentsps1) | Compliance policies incl. actions for noncompliance | ✅ Available |

> **CSV format:** UTF-8 with BOM, delimiter `;` by default (opens directly in Excel with German/European regional settings). Use `-Delimiter ','` for English regional settings.

---

## Get-IntuneConfigAssignments.ps1

### What it covers

| Policy type | Graph endpoint (beta) |
|---|---|
| Templates (classic device configurations) | `/deviceManagement/deviceConfigurations` |
| Settings Catalog, Endpoint Security, Security Baselines | `/deviceManagement/configurationPolicies` |
| Administrative Templates (ADMX) | `/deviceManagement/groupPolicyConfigurations` |
| Endpoint Security (legacy intents) | `/deviceManagement/intents` |

### Features

- **One row per policy** – all assignments listed in the `AssignedGroups` column, include before exclude, filters inline:
  ```
  GRP-Win-Pilot | GRP-Win-Prod [Filter: Laptops only (include)] | Exclude: GRP-Win-Kiosk
  ```
- Optional **one row per assignment** (`-OneRowPerAssignment`) for pivot tables in Excel
- Resolves group names and assignment filter names; flags **deleted groups and filters**
- **"Which policies hit this group?"** via `-GroupName` (wildcards supported), optionally including policies assigned to *All devices* / *All users* (`-IncludeTenantWide`)
- Finds **unassigned policies** (`-IncludeUnassigned`)
- Group names and legacy intent assignments resolved via **JSON batching** (`$batch`) with retry on throttling (429/503/504)
- **Fault-tolerant:** if one policy type cannot be read, the report still completes and is clearly marked as incomplete
- Works with **delegated** and **app-only** authentication; `-TenantId` for multi-tenant admins

### Requirements

- PowerShell 5.1 or 7.x
- Module `Microsoft.Graph.Authentication`
  ```powershell
  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
  ```
- Graph permissions (delegated or application):

  | Permission | Also accepted |
  |---|---|
  | `DeviceManagementConfiguration.Read.All` | `DeviceManagementConfiguration.ReadWrite.All` |
  | `Group.Read.All` | `Group.ReadWrite.All`, `Directory.Read.All` |

### Usage

```powershell
# Full report (interactive sign-in if not connected)
.\Get-IntuneConfigAssignments.ps1 -TenantId contoso.onmicrosoft.com

# Which policies are assigned to these groups? Include tenant-wide assignments
.\Get-IntuneConfigAssignments.ps1 -GroupName "GRP-Intune-Win-*" -IncludeTenantWide -GridView

# Include policies without any assignment, custom export path
.\Get-IntuneConfigAssignments.ps1 -IncludeUnassigned -ExportPath C:\Reports\assignments.csv

# One row per assignment (raw data)
.\Get-IntuneConfigAssignments.ps1 -OneRowPerAssignment

# Use the objects in the pipeline
$r = .\Get-IntuneConfigAssignments.ps1 -PassThru
$r | Where-Object AssignmentCount -gt 1

# Comma-separated CSV (English regional settings)
.\Get-IntuneConfigAssignments.ps1 -Delimiter ','
```

#### App-only (automation)

Connect before running the script – it reuses the existing connection and verifies the application permissions with test requests:

```powershell
Connect-MgGraph -TenantId <tenant> -ClientId <appId> -CertificateThumbprint <thumbprint>
.\Get-IntuneConfigAssignments.ps1
```

#### Sign out

```powershell
Disconnect-MgGraph
```

### Parameters

| Parameter | Description |
|---|---|
| `-TenantId` | Tenant ID or primary domain. Reconnects if the current session belongs to another tenant. |
| `-GroupName` | Only policies with a direct assignment to a matching group (wildcards). Shows **all** assignments of these policies. Nested group membership is not resolved. |
| `-IncludeTenantWide` | With `-GroupName`: also include policies assigned to *All devices* / *All users*. |
| `-IncludeUnassigned` | Also list policies without assignments (no effect with `-GroupName`). |
| `-OneRowPerAssignment` | One row per assignment instead of per policy, with separate columns for mode and filter. |
| `-ExportPath` | CSV path. Default: `%TEMP%\Intune-ConfigAssignments_<timestamp>.csv`. Folder is created if missing. |
| `-Delimiter` | CSV delimiter: `;` (default), `,` or tab. |
| `-GridView` | Additionally show the result in `Out-GridView`. |
| `-PassThru` | Return the result objects to the pipeline. |

### Output columns (default mode)

| Column | Content |
|---|---|
| `Policy` | Policy name |
| `Type` | Policy type (Template, Settings Catalog, Endpoint Security / Baseline, Administrative Template, Endpoint Security (legacy intent)) |
| `Platform` | Normalized platform (Windows, iOS, macOS, Android, Linux) |
| `DetailType` | OData type, technologies or template family |
| `AssignmentCount` | Number of assignments (include, exclude, tenant-wide) |
| `AssignedGroups` | All assignments, separated by ` \| ` |
| `GroupIds` | Group IDs in the same order as `AssignedGroups` |
| `PolicyId` | Policy ID |

With `-OneRowPerAssignment` the columns are: `Policy`, `Type`, `Platform`, `DetailType`, `Mode` (Include/Exclude), `Scope` (Group/TenantWide/None/Unknown), `Group`, `Filter`, `FilterMode`, `GroupId`, `PolicyId`.

### Limitations

- Only **direct** assignments are evaluated – nested group membership is not resolved. For the effective state of a single device use *Intune > Devices > [device] > Device configuration*.
- Uses Graph **beta** endpoints (no complete v1.0 equivalent exists for Settings Catalog and ADMX policies).
- Compliance policies, apps, scripts and remediations are not part of this script.

---

## Get-IntuneComplianceAssignments.ps1

Same approach, parameters and output format as `Get-IntuneConfigAssignments.ps1`, plus the **actions for noncompliance** of every policy and the **tenant-wide compliance settings**.

### What it covers

| Policy type | Graph endpoint (beta) |
|---|---|
| Compliance policies (classic) | `/deviceManagement/deviceCompliancePolicies` |
| Compliance policies (Settings Catalog based, e.g. Linux) | `/deviceManagement/compliancePolicies` |
| Actions for noncompliance | `…/{id}/scheduledActionsForRule?$expand=scheduledActionConfigurations` |
| Tenant-wide compliance settings | `/deviceManagement?$select=settings` |

### Additional output

- **Console:** tenant-wide settings
  - *Mark devices with no compliance policy assigned as* (compliant / not compliant)
  - *Compliance status validity period* (days)
- **Column `NoncompliantAfter`:** when a device is marked as noncompliant (`immediately`, `1 day`, `3 days`, …)
- **Column `Actions`:** all actions for noncompliance in chronological order, including notification template and CC groups:
  ```
  Mark device noncompliant: immediately | Send email to end user: after 1 day (Template: Compliance notice; CC: GRP-IT-Helpdesk) | Remotely lock the noncompliant device: after 3 days
  ```
  The console table omits `Actions` for space reasons; it is included in the CSV, `-GridView` and `-PassThru`.

### Permissions

Same as above (`DeviceManagementConfiguration.Read.All`, `Group.Read.All`).
Optional: `DeviceManagementServiceConfig.Read.All` – needed to resolve **notification template names** and read the **tenant-wide settings**. Without it the report still runs; template IDs are shown instead of names and the tenant settings are skipped with a warning.

### Usage

```powershell
.\Get-IntuneComplianceAssignments.ps1 -TenantId contoso.onmicrosoft.com
.\Get-IntuneComplianceAssignments.ps1 -GroupName "GRP-Intune-Win-*" -IncludeTenantWide
.\Get-IntuneComplianceAssignments.ps1 -IncludeUnassigned -OneRowPerAssignment
```

Default export path: `%TEMP%\Intune-ComplianceAssignments_<timestamp>.csv`

### Limitations

- The built-in *Default Device Compliance Policy* is not exposed via Graph and therefore not listed.
- Only direct assignments are evaluated (no nested groups).

---

## Disclaimer

The scripts are read-only and make no changes to your tenant. Nevertheless, test them in a non-production tenant first. Provided as-is, without warranty.
