#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Tears down an Exchange Online RBAC-for-Applications scoping created by the
    RBAC Builder: role assignment(s), management scope, EXO service principal,
    and the access security group.

.DESCRIPTION
    Given the app display name that was used when the RBAC was built, this
    discovers and removes (in dependency order):
        1. Management Role Assignment(s) for the app's EXO service principal
        2. The Management Scope(s) those assignments referenced
        3. The EXO Service Principal
        4. The mail-enabled security access group ("<App>_RBAC_Access")

    Discovery is by naming convention + role-assignee match, so only the app
    name is needed (no Graph connection required). Everything found is previewed
    and the operator must type the app name to confirm before anything is
    removed. Results are written to C:\MSP-M365-Utility\.

.NOTES
    Required Modules:
        - ExchangeOnlineManagement
    Required Permissions:
        - Exchange Online : Organization Management
#>

# ----------------------------------------------------------
#  CONFIGURATION
# ----------------------------------------------------------
$ErrorActionPreference = 'Stop'
$Host.UI.RawUI.WindowTitle = "M365 RBAC Removal"

# ----------------------------------------------------------
#  BANNER
# ----------------------------------------------------------
Clear-Host
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |          M365 RBAC-for-Applications Removal      |" -ForegroundColor Cyan
Write-Host "  |              Exchange Online                     |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""

# ----------------------------------------------------------
#  INPUT - TENANT CODE
# ----------------------------------------------------------
do {
    $TenantCode = (Read-Host "  Enter the three-letter Tenant Code (e.g. ABC)").Trim().ToUpper()
    if ($TenantCode -notmatch '^[A-Z]{3}$') {
        Write-Host "  [!] Invalid input. Please enter exactly 3 letters (A-Z)." -ForegroundColor Yellow
    }
} while ($TenantCode -notmatch '^[A-Z]{3}$')

# ----------------------------------------------------------
#  OUTPUT PATH
# ----------------------------------------------------------
$OutputRoot = 'C:\MSP-M365-Utility'
if (-not (Test-Path $OutputRoot)) {
    try { New-Item -ItemType Directory -Path $OutputRoot -Force -ErrorAction Stop | Out-Null }
    catch { Write-Host "  [ERROR] Could not create '$OutputRoot': $_" -ForegroundColor Red; Read-Host "`n  Press Enter to exit"; exit 1 }
}
$Timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputPath = Join-Path $OutputRoot "RBACRemoval_${TenantCode}_${Timestamp}.csv"

# ----------------------------------------------------------
#  RESULT LOGGING HELPER
# ----------------------------------------------------------
$Results = [System.Collections.Generic.List[PSCustomObject]]::new()
function Add-Result {
    param([string]$Item, [string]$Name, [string]$Status, [string]$Detail = '')
    $Results.Add([PSCustomObject]@{ Item = $Item; Name = $Name; Status = $Status; Detail = $Detail })
    $col = switch ($Status) {
        'Removed'   { 'Green' }
        'Not-Found' { 'DarkGray' }
        'Info'      { 'Cyan' }
        default     { 'Red' }
    }
    Write-Host ("  {0,-16} {1,-42} {2,-10} {3}" -f $Item, $Name, $Status, $Detail) -ForegroundColor $col
}

# ----------------------------------------------------------
#  MODULE CHECK + CONNECT
# ----------------------------------------------------------
Write-Host "  [1/3] Checking required modules..." -ForegroundColor Cyan
foreach ($Mod in @('ExchangeOnlineManagement')) {
    if (-not (Get-Module -ListAvailable -Name $Mod)) {
        Write-Host "  [!] Module '$Mod' not found. Installing..." -ForegroundColor Yellow
        Install-Module -Name $Mod -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module -Name $Mod -ErrorAction Stop
    Write-Host "  [OK] $Mod loaded." -ForegroundColor Green
}

Write-Host ""
Write-Host "  [2/3] Connecting to Exchange Online..." -ForegroundColor Cyan
try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "  [OK] Exchange Online connected." -ForegroundColor Green
} catch {
    Write-Host "  [ERROR] Failed to connect: $_" -ForegroundColor Red
    Read-Host "`n  Press Enter to exit"; exit 1
}

# ----------------------------------------------------------
#  DISCOVER OBJECTS
# ----------------------------------------------------------
$AppDisplayName = (Read-Host "  App display name (as used when the RBAC was built)").Trim()
$ExoSpName = "$($AppDisplayName)_Service_Principal"
$GroupName = "$($AppDisplayName)_RBAC_Access"

Write-Host ""
Write-Host "  Discovering objects..." -ForegroundColor Cyan

# EXO service principal (by convention name, then by any AppId match on that name)
$exoSp = $null
try { $exoSp = Get-ServicePrincipal -Identity $ExoSpName -ErrorAction SilentlyContinue } catch { }
if (-not $exoSp) {
    try { $exoSp = Get-ServicePrincipal -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -eq $ExoSpName } | Select-Object -First 1 } catch { }
}

# Role assignments for that SP
$roleAsgns = @()
$scopeNames = @()
if ($exoSp) {
    try {
        $roleAsgns = @(Get-ManagementRoleAssignment -ErrorAction SilentlyContinue |
            Where-Object { $_.RoleAssigneeName -eq $exoSp.DisplayName -or $_.App -eq $exoSp.ObjectId -or $_.RoleAssignee -eq $exoSp.ObjectId })
        $scopeNames = @($roleAsgns | ForEach-Object { $_.CustomResourceScope } | Where-Object { $_ } | Select-Object -Unique)
    } catch { }
}

# Access group
$grp = $null
try { $grp = Get-DistributionGroup -Identity $GroupName -ErrorAction SilentlyContinue } catch { }

# If no scope discovered from assignments, offer a manual name
if ($scopeNames.Count -eq 0) {
    $manualScope = (Read-Host "  Management Scope name to remove (blank to skip)").Trim()
    if ($manualScope) { $scopeNames = @($manualScope) }
}

# ----------------------------------------------------------
#  PREVIEW
# ----------------------------------------------------------
Write-Host ""
Write-Host "  The following will be REMOVED:" -ForegroundColor Yellow
Write-Host ("    Role Assignments : {0}" -f $(if ($roleAsgns.Count) { ($roleAsgns.Name -join ', ') } else { '(none found)' }))
Write-Host ("    Management Scope  : {0}" -f $(if ($scopeNames.Count) { ($scopeNames -join ', ') } else { '(none found)' }))
Write-Host ("    Service Principal : {0}" -f $(if ($exoSp) { $exoSp.DisplayName } else { '(none found)' }))
Write-Host ("    Access Group      : {0}" -f $(if ($grp) { $grp.Name } else { '(none found)' }))

if (-not $roleAsgns.Count -and -not $scopeNames.Count -and -not $exoSp -and -not $grp) {
    Write-Host ""
    Write-Host "  [!] Nothing found for '$AppDisplayName'. Check the app name and try again." -ForegroundColor Yellow
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Read-Host "`n  Press Enter to exit"; exit 0
}

Write-Host ""
Write-Host "  This is a full teardown and cannot be undone." -ForegroundColor Red
$typed = Read-Host "  To confirm, type the app name exactly ($AppDisplayName)"
if ($typed.Trim() -ne $AppDisplayName) {
    Write-Host "  Cancelled - input did not match." -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Read-Host "`n  Press Enter to exit"; exit 0
}

$ScriptStart = [System.Diagnostics.Stopwatch]::StartNew()
Write-Host ""
Write-Host "  [3/3] Removing (role assignments -> scope -> service principal -> group)..." -ForegroundColor Cyan
Write-Host ""

# 1) Role assignments
foreach ($ra in $roleAsgns) {
    try {
        Remove-ManagementRoleAssignment -Identity $ra.Identity -Confirm:$false -ErrorAction Stop
        Add-Result 'RoleAssignment' $ra.Name 'Removed' ''
    } catch { Add-Result 'RoleAssignment' $ra.Name 'Failed' $_.Exception.Message }
}
if ($roleAsgns.Count -eq 0) { Add-Result 'RoleAssignment' '-' 'Not-Found' '' }

# 2) Management scope(s)
foreach ($sn in $scopeNames) {
    $sc = $null
    try { $sc = Get-ManagementScope -Identity $sn -ErrorAction SilentlyContinue } catch { }
    if ($sc) {
        try {
            Remove-ManagementScope -Identity $sn -Confirm:$false -ErrorAction Stop
            Add-Result 'Scope' $sn 'Removed' ''
        } catch { Add-Result 'Scope' $sn 'Failed' $_.Exception.Message }
    } else {
        Add-Result 'Scope' $sn 'Not-Found' ''
    }
}
if ($scopeNames.Count -eq 0) { Add-Result 'Scope' '-' 'Not-Found' '' }

# 3) EXO service principal
if ($exoSp) {
    try {
        Remove-ServicePrincipal -Identity $exoSp.ObjectId -Confirm:$false -ErrorAction Stop
        Add-Result 'ServicePrincipal' $exoSp.DisplayName 'Removed' ''
    } catch { Add-Result 'ServicePrincipal' $exoSp.DisplayName 'Failed' $_.Exception.Message }
} else {
    Add-Result 'ServicePrincipal' $ExoSpName 'Not-Found' ''
}

# 4) Access group
if ($grp) {
    try {
        Remove-DistributionGroup -Identity $grp.Identity -Confirm:$false -ErrorAction Stop
        Add-Result 'Group' $grp.Name 'Removed' ''
    } catch { Add-Result 'Group' $grp.Name 'Failed' $_.Exception.Message }
} else {
    Add-Result 'Group' $GroupName 'Not-Found' ''
}

# ----------------------------------------------------------
#  EXPORT + DISCONNECT + SUMMARY
# ----------------------------------------------------------
try {
    $Results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "  [OK] Results CSV exported." -ForegroundColor Green
    Write-Host "       Path: $OutputPath" -ForegroundColor DarkGray
} catch {
    Write-Host "  [ERROR] Failed to export results: $_" -ForegroundColor Red
}

Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Write-Host "  Session disconnected." -ForegroundColor DarkGray

$ScriptStart.Stop()
$Elapsed = $ScriptStart.Elapsed
$RemovedCount = ($Results | Where-Object { $_.Status -eq 'Removed' }).Count
$FailCount    = ($Results | Where-Object { $_.Status -eq 'Failed'  }).Count
$RunTime = "{0:D2}h {1:D2}m {2:D2}s {3:D3}ms" -f $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds, $Elapsed.Milliseconds

Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |                  RUN SUMMARY                    |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Tenant Code       : {0,-29}|" -f $TenantCode)     -ForegroundColor White
Write-Host ("  | App               : {0,-29}|" -f $AppDisplayName) -ForegroundColor White
Write-Host ("  | Objects Removed   : {0,-29}|" -f $RemovedCount)   -ForegroundColor White
Write-Host ("  | Failed            : {0,-29}|" -f $FailCount)      -ForegroundColor White
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Total Run Time    : {0,-29}|" -f $RunTime)        -ForegroundColor Yellow
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""
Read-Host "  Press Enter to exit"
