#Requires -Modules Microsoft.Graph.Applications, ExchangeOnlineManagement

<#
.SYNOPSIS
    Provisions Exchange Online RBAC-for-Applications so a Graph app
    (Application permissions) is scoped to a mail-enabled security group of
    mailboxes instead of the whole tenant.

.DESCRIPTION
    The security group is the durable, SCALABLE access boundary: the management
    scope filters on membership of that group, so granting a new mailbox access
    later is just "add it to the group" - no need to re-run this builder.

    Flow (idempotent - existing objects are reused, not duplicated):
        1. Connect Microsoft Graph (first), then Exchange Online
        2. Resolve the app from a typed display name -> AppId / ObjectId
        3. Register the app as an EXO Service Principal
        4. Create (or reuse) a mail-enabled security group and seed the pasted
           mailboxes into it
        5. Create (or reuse) a Management Scope filtered on group membership
        6. Map the app's consented Graph permissions to EXO "Application X" roles
        7. Create (or reuse) the Management Role Assignment(s)
        8. Print the EXO Service Principal Object ID and the group to grow later
        9. Optional test loop: Test-ServicePrincipalAuthorization per address

    Results are written to C:\MSP-M365-Utility\.

.NOTES
    Required Modules:
        - Microsoft.Graph.Applications
        - ExchangeOnlineManagement
    Required Permissions:
        - Microsoft Graph : Application.Read.All
        - Exchange Online : Organization Management (RBAC + recipient changes)
#>

# ----------------------------------------------------------
#  CONFIGURATION
# ----------------------------------------------------------
$ErrorActionPreference = 'Stop'
$Host.UI.RawUI.WindowTitle = "M365 RBAC Builder"

# ----------------------------------------------------------
#  BANNER
# ----------------------------------------------------------
Clear-Host
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |          M365 RBAC-for-Applications Builder      |" -ForegroundColor Cyan
Write-Host "  |        Microsoft Graph  *  Exchange Online       |" -ForegroundColor Cyan
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
$OutputPath = Join-Path $OutputRoot "RBACBuilder_${TenantCode}_${Timestamp}.csv"

# ----------------------------------------------------------
#  RESULT LOGGING HELPER
# ----------------------------------------------------------
$Results = [System.Collections.Generic.List[PSCustomObject]]::new()
function Add-Result {
    param([string]$Item, [string]$Name, [string]$Status, [string]$Detail = '')
    $Results.Add([PSCustomObject]@{ Item = $Item; Name = $Name; Status = $Status; Detail = $Detail })
    $col = switch ($Status) {
        'Created'  { 'Green' }
        'Added'    { 'Green' }
        'Reused'   { 'DarkCyan' }
        'Skipped'  { 'DarkGray' }
        'Info'     { 'Cyan' }
        default    { 'Red' }   # Failed
    }
    Write-Host ("  {0,-14} {1,-40} {2,-9} {3}" -f $Item, $Name, $Status, $Detail) -ForegroundColor $col
}

# ----------------------------------------------------------
#  MODULE CHECK + CONNECT (Graph first, then EXO)
# ----------------------------------------------------------
Write-Host "  [1/4] Checking required modules..." -ForegroundColor Cyan
foreach ($Mod in @('Microsoft.Graph.Applications','ExchangeOnlineManagement')) {
    if (-not (Get-Module -ListAvailable -Name $Mod)) {
        Write-Host "  [!] Module '$Mod' not found. Installing..." -ForegroundColor Yellow
        Install-Module -Name $Mod -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module -Name $Mod -ErrorAction Stop
    Write-Host "  [OK] $Mod loaded." -ForegroundColor Green
}

Write-Host ""
Write-Host "  [2/4] Connecting (Graph first, then Exchange Online)..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes "Application.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "  [OK] Microsoft Graph connected." -ForegroundColor Green
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "  [OK] Exchange Online connected." -ForegroundColor Green
} catch {
    Write-Host "  [ERROR] Failed to connect: $_" -ForegroundColor Red
    Read-Host "`n  Press Enter to exit"; exit 1
}

# ----------------------------------------------------------
#  RESOLVE APP
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [3/4] Resolving app and collecting inputs..." -ForegroundColor Cyan
$AppDisplayName = (Read-Host "  App display name (Entra App Registration)").Trim()

try {
    $spMatches = @(Get-MgServicePrincipal -Filter "displayName eq '$($AppDisplayName.Replace("'","''"))'" -ErrorAction Stop)
} catch {
    Write-Host "  [ERROR] Graph lookup failed: $_" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Read-Host "`n  Press Enter to exit"; exit 1
}

if ($spMatches.Count -eq 0) {
    Write-Host "  [ERROR] No Entra service principal found for '$AppDisplayName'." -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Read-Host "`n  Press Enter to exit"; exit 1
}
if ($spMatches.Count -gt 1) {
    Write-Host "  Multiple apps matched '$AppDisplayName':" -ForegroundColor Yellow
    for ($i = 0; $i -lt $spMatches.Count; $i++) {
        Write-Host ("    [{0}] {1}  AppId: {2}" -f $i, $spMatches[$i].DisplayName, $spMatches[$i].AppId)
    }
    do { $sel = Read-Host "  Enter the number of the correct app" } while ($sel -notmatch '^\d+$' -or [int]$sel -ge $spMatches.Count)
    $EntraSp = $spMatches[[int]$sel]
} else {
    $EntraSp = $spMatches[0]
}
$EntraAppId    = $EntraSp.AppId
$EntraObjectId = $EntraSp.Id
Write-Host "  [OK] App: $($EntraSp.DisplayName)  AppId: $EntraAppId" -ForegroundColor Green

# ----------------------------------------------------------
#  COLLECT NAMES + SEED MAILBOXES
# ----------------------------------------------------------
$FriendlyName = (Read-Host "  Friendly name for this access set (used to name the scope/assignment)").Trim()
if (-not $FriendlyName) { $FriendlyName = $AppDisplayName }

$ExoSpName    = "$($AppDisplayName)_Service_Principal"
$GroupName    = "$($AppDisplayName)_RBAC_Access"
$GroupAlias   = ($GroupName -replace '[^a-zA-Z0-9]', '')
$ScopeName    = "$($FriendlyName)_Management_Scope"
$RoleAsgnBase = "$($FriendlyName)_RoleAssignment"

$EmailRegex = '^[^@\s]+@[^@\s]+\.[^@\s]+$'
Write-Host ""
Write-Host "  Paste the mailbox SMTP addresses to seed into the access group" -ForegroundColor Yellow
Write-Host "  (one per line). You can add more later by adding members to the group." -ForegroundColor DarkGray
Write-Host "  Blank line to finish:" -ForegroundColor Yellow
$Mailboxes = [System.Collections.Generic.List[string]]::new()
$seen = @{}
while ($true) {
    $line = Read-Host
    if ($line -eq '') { break }
    foreach ($cand in ($line -split '[,;\s]+')) {
        $e = $cand.Trim()
        if (-not $e) { continue }
        if ($e -notmatch $EmailRegex) { Write-Host "    [!] Skipped invalid: $e" -ForegroundColor DarkGray; continue }
        $k = $e.ToLower(); if ($seen.ContainsKey($k)) { continue }; $seen[$k] = $true
        [void]$Mailboxes.Add($e)
    }
}

# ----------------------------------------------------------
#  PREVIEW + CONFIRM
# ----------------------------------------------------------
Write-Host ""
Write-Host "  Will provision:" -ForegroundColor Cyan
Write-Host "    App                 : $AppDisplayName ($EntraAppId)"
Write-Host "    EXO Svc Principal   : $ExoSpName"
Write-Host "    Access Group        : $GroupName  ($($Mailboxes.Count) seed member(s))"
Write-Host "    Management Scope    : $ScopeName"
Write-Host "    Role Assignment base: $RoleAsgnBase"
$ok = Read-Host "`n  Proceed? (Y/N)"
if ($ok -notin @('Y','y')) {
    Write-Host "  Aborted." -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Read-Host "`n  Press Enter to exit"; exit 0
}

$ScriptStart = [System.Diagnostics.Stopwatch]::StartNew()
Write-Host ""
Write-Host "  [4/4] Provisioning..." -ForegroundColor Cyan
Write-Host ""

# ----- EXO Service Principal (idempotent) -----
$exoSp = $null
try {
    $exoSp = Get-ServicePrincipal -ErrorAction Stop | Where-Object { $_.AppId -eq $EntraAppId } | Select-Object -First 1
} catch { }
if ($exoSp) {
    Add-Result 'ServicePrincipal' $exoSp.DisplayName 'Reused' "ObjectId: $($exoSp.ObjectId)"
} else {
    try {
        $exoSp = New-ServicePrincipal -AppId $EntraAppId -ObjectId $EntraObjectId -DisplayName $ExoSpName -ErrorAction Stop
        Add-Result 'ServicePrincipal' $ExoSpName 'Created' "ObjectId: $($exoSp.ObjectId)"
    } catch {
        Add-Result 'ServicePrincipal' $ExoSpName 'Failed' $_.Exception.Message
    }
}
if (-not $exoSp) {
    $Results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "  [ERROR] Cannot continue without an EXO service principal." -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Read-Host "`n  Press Enter to exit"; exit 1
}
$ExoSpObjectId = $exoSp.ObjectId

# ----- Access group (idempotent) -----
$grp = $null
try { $grp = Get-DistributionGroup -Identity $GroupName -ErrorAction SilentlyContinue } catch { }
if ($grp) {
    Add-Result 'Group' $GroupName 'Reused' 'Already existed'
} else {
    try {
        $grp = New-DistributionGroup -Name $GroupName -Alias $GroupAlias -Type Security -ErrorAction Stop
        Add-Result 'Group' $GroupName 'Created' 'Mail-enabled security group'
        Start-Sleep -Seconds 5   # let the new group replicate before adding members
    } catch {
        Add-Result 'Group' $GroupName 'Failed' $_.Exception.Message
    }
}

# ----- Seed members -----
if ($grp) {
    foreach ($mbx in $Mailboxes) {
        try {
            Add-DistributionGroupMember -Identity $GroupName -Member $mbx -ErrorAction Stop
            Add-Result 'Member' $mbx 'Added' ''
        } catch {
            if ("$($_.Exception.Message)" -match 'already a member') {
                Add-Result 'Member' $mbx 'Skipped' 'Already a member'
            } else {
                Add-Result 'Member' $mbx 'Failed' $_.Exception.Message
            }
        }
    }
}

# ----- Management scope (idempotent) -----
$scopeOk = $false
if ($grp) {
    $existingScope = $null
    try { $existingScope = Get-ManagementScope -Identity $ScopeName -ErrorAction SilentlyContinue } catch { }
    if ($existingScope) {
        Add-Result 'Scope' $ScopeName 'Reused' 'Already existed'
        $scopeOk = $true
    } else {
        try {
            $groupDn = (Get-DistributionGroup -Identity $GroupName -ErrorAction Stop).DistinguishedName
            New-ManagementScope -Name $ScopeName -RecipientRestrictionFilter "MemberOfGroup -eq '$groupDn'" -ErrorAction Stop | Out-Null
            Add-Result 'Scope' $ScopeName 'Created' 'Filtered on group membership'
            $scopeOk = $true
        } catch {
            Add-Result 'Scope' $ScopeName 'Failed' $_.Exception.Message
        }
    }
}

# ----- Map consented Graph permissions -> EXO roles -----
$RoleMap = @{
    'Mail.Read'                = 'Application Mail.Read'
    'Mail.ReadWrite'           = 'Application Mail.ReadWrite'
    'Mail.Send'                = 'Application Mail.Send'
    'MailboxSettings.Read'     = 'Application MailboxSettings.Read'
    'MailboxSettings.ReadWrite'= 'Application MailboxSettings.ReadWrite'
    'Calendars.Read'           = 'Application Calendars.Read'
    'Calendars.ReadWrite'      = 'Application Calendars.ReadWrite'
    'Contacts.Read'            = 'Application Contacts.Read'
    'Contacts.ReadWrite'       = 'Application Contacts.ReadWrite'
    'EWS.AccessAsApp'          = 'Application EWS.AccessAsApp'
}
$MappedRoles = @()
try {
    $graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -ErrorAction Stop
    $assignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $EntraObjectId -ErrorAction Stop |
        Where-Object { $_.ResourceId -eq $graphSp.Id }
    $consented = foreach ($a in $assignments) { ($graphSp.AppRoles | Where-Object { $_.Id -eq $a.AppRoleId }).Value }
    foreach ($perm in $consented) {
        if ($RoleMap.ContainsKey($perm)) { $MappedRoles += $RoleMap[$perm] }
        else { Add-Result 'Permission' $perm 'Skipped' 'No matching EXO Application role' }
    }
    $MappedRoles = @($MappedRoles | Select-Object -Unique)
    Add-Result 'Permissions' "$($consented.Count) consented" 'Info' "Mapped to $($MappedRoles.Count) EXO role(s)"
} catch {
    Add-Result 'Permissions' '-' 'Failed' "Could not read app role assignments: $($_.Exception.Message)"
}

# ----- Role assignments (idempotent) -----
if ($scopeOk -and $MappedRoles.Count -gt 0) {
    foreach ($role in $MappedRoles) {
        $asgnName = "$RoleAsgnBase-$($role -replace '\s','')"
        $existingAsgn = $null
        try { $existingAsgn = Get-ManagementRoleAssignment -Identity $asgnName -ErrorAction SilentlyContinue } catch { }
        if ($existingAsgn) {
            Add-Result 'RoleAssignment' $asgnName 'Reused' $role
        } else {
            try {
                New-ManagementRoleAssignment -Name $asgnName -App $ExoSpObjectId -Role $role -CustomResourceScope $ScopeName -ErrorAction Stop | Out-Null
                Add-Result 'RoleAssignment' $asgnName 'Created' $role
            } catch {
                Add-Result 'RoleAssignment' $asgnName 'Failed' $_.Exception.Message
            }
        }
    }
} elseif ($MappedRoles.Count -eq 0) {
    Add-Result 'RoleAssignment' '-' 'Skipped' 'No mapped roles to assign'
}

# ----------------------------------------------------------
#  SCALABILITY NOTE + OPTIONAL TEST LOOP
# ----------------------------------------------------------
Write-Host ""
Write-Host "  EXO Service Principal Object ID : $ExoSpObjectId" -ForegroundColor Cyan
Write-Host "  To grant MORE mailboxes access later, just add them to the group:" -ForegroundColor Cyan
Write-Host "      $GroupName" -ForegroundColor White
Write-Host "  (e.g. via the Add Distribution List Members module) - no re-run needed." -ForegroundColor DarkGray

$runTest = Read-Host "`n  Run the scope test loop now? (Y/N)"
if ($runTest -in @('Y','y')) {
    Write-Host "  Type an email to test, or 'done' to finish." -ForegroundColor Cyan
    while ($true) {
        $testEmail = (Read-Host "  Email to test").Trim()
        if ($testEmail -eq 'done' -or $testEmail -eq '') { break }
        try {
            $tr = Test-ServicePrincipalAuthorization -Identity $ExoSpObjectId -Resource $testEmail -ErrorAction Stop
            $inScope = [bool]($tr | Where-Object { $_.Granted -eq $true })
            $col = if ($inScope) { 'Green' } else { 'Red' }
            Write-Host ("    $testEmail -> InScope: $inScope") -ForegroundColor $col
            Add-Result 'Test' $testEmail $(if ($inScope) { 'Info' } else { 'Failed' }) "InScope: $inScope"
        } catch {
            Write-Host "    [ERROR] $($_.Exception.Message)" -ForegroundColor Red
            Add-Result 'Test' $testEmail 'Failed' $_.Exception.Message
        }
    }
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
Disconnect-MgGraph -ErrorAction SilentlyContinue
Write-Host "  Sessions disconnected." -ForegroundColor DarkGray

$ScriptStart.Stop()
$Elapsed = $ScriptStart.Elapsed
$FailCount = ($Results | Where-Object { $_.Status -eq 'Failed' }).Count
$RunTime = "{0:D2}h {1:D2}m {2:D2}s {3:D3}ms" -f $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds, $Elapsed.Milliseconds

Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |                  RUN SUMMARY                    |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Tenant Code       : {0,-29}|" -f $TenantCode)   -ForegroundColor White
Write-Host ("  | App               : {0,-29}|" -f $AppDisplayName) -ForegroundColor White
Write-Host ("  | Access Group      : {0,-29}|" -f $GroupName)    -ForegroundColor White
Write-Host ("  | Seed Members      : {0,-29}|" -f $Mailboxes.Count) -ForegroundColor White
Write-Host ("  | Failed Actions    : {0,-29}|" -f $FailCount)    -ForegroundColor White
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Total Run Time    : {0,-29}|" -f $RunTime)      -ForegroundColor Yellow
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""
Read-Host "  Press Enter to exit"
