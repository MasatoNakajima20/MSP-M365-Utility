#Requires -Modules ExchangeOnlineManagement, Microsoft.Graph.Users, Microsoft.Graph.Groups

<#
.SYNOPSIS
    Discover a user's mailbox access and group memberships, then selectively
    remove what you choose.

.DESCRIPTION
    Given a user's email, this script enumerates:

      1. Mailbox access:
         - FullAccess permissions on other mailboxes (FullAccess scan scope is
           selectable: SharedMailbox only, or all mailboxes).
         - SendAs permissions across the org.
      2. Group memberships:
         - Distribution Groups
         - Mail-Enabled Security Groups
         - Microsoft 365 Groups
         - Pure Security Groups are listed for visibility but are NOT eligible
           for removal here (this is a "distribution membership" cleanup).

    The script lists everything it finds, then prompts twice (once for the
    mailbox set, once for the group set). Each prompt accepts a multiline
    paste:

        (blank line straight away)  - skip that phase
        all                          - target every discovered item in that phase
        <list of emails / group SMTP addresses, one per line>

    Each removal has a Y/N confirmation. Results are written to
    C:\MSP-M365-Utility\.

.NOTES
    Required Modules:
        - ExchangeOnlineManagement
        - Microsoft.Graph.Users
        - Microsoft.Graph.Groups
    Required Permissions:
        - Exchange Online : Recipient Management (or higher)
        - Microsoft Graph : User.Read.All, Group.Read.All, GroupMember.ReadWrite.All
#>

# ----------------------------------------------------------
#  CONFIGURATION
# ----------------------------------------------------------
$ErrorActionPreference = 'Stop'
$Host.UI.RawUI.WindowTitle = "M365 User Access - Discover and Remove"

# ----------------------------------------------------------
#  BANNER
# ----------------------------------------------------------
Clear-Host
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |     M365 User Access - Discover and Remove       |" -ForegroundColor Cyan
Write-Host "  |   Exchange Online  *  Microsoft Graph            |" -ForegroundColor Cyan
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
#  INPUT - TARGET USER
# ----------------------------------------------------------
$EmailRegex = '^[^@\s]+@[^@\s]+\.[^@\s]+$'
do {
    $TargetUser = (Read-Host "  Enter the USER email to audit").Trim()
    if ($TargetUser -notmatch $EmailRegex) {
        Write-Host "  [!] That doesn't look like an email address." -ForegroundColor Yellow
    }
} while ($TargetUser -notmatch $EmailRegex)

# ----------------------------------------------------------
#  INPUT - FULLACCESS SCAN SCOPE
# ----------------------------------------------------------
$ScopeAnswer = (Read-Host "  FullAccess scan scope - [S]hared only / [A]ll mailboxes  [A]").Trim().ToUpper()
$Scope = if ($ScopeAnswer -eq 'S') { 'Shared' } else { 'All' }

# ----------------------------------------------------------
#  OUTPUT PATH
# ----------------------------------------------------------
$OutputRoot = 'C:\MSP-M365-Utility'
if (-not (Test-Path $OutputRoot)) {
    try {
        New-Item -ItemType Directory -Path $OutputRoot -Force -ErrorAction Stop | Out-Null
    } catch {
        Write-Host "  [ERROR] Could not create '$OutputRoot': $_" -ForegroundColor Red
        Read-Host "`n  Press Enter to exit"; exit 1
    }
}
$Timestamp     = Get-Date -Format "yyyyMMdd_HHmmss"
$AuditCsv      = Join-Path $OutputRoot "UserAccessAudit_${TenantCode}_${Timestamp}.csv"
$RemovalCsv    = Join-Path $OutputRoot "UserAccessRemoval_${TenantCode}_${Timestamp}.csv"

# ----------------------------------------------------------
#  MODULE CHECK + CONNECT
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [1/5] Checking required modules..." -ForegroundColor Cyan
# Graph modules first so their newer MSAL is available before EXO's older
# bundled copy. (Connect order is Graph-first below, on all PS versions.)
foreach ($Mod in @('Microsoft.Graph.Users','Microsoft.Graph.Groups','ExchangeOnlineManagement')) {
    if (-not (Get-Module -ListAvailable -Name $Mod)) {
        Write-Host "  [!] Module '$Mod' not found. Installing..." -ForegroundColor Yellow
        Install-Module -Name $Mod -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module -Name $Mod -ErrorAction Stop
    Write-Host "  [OK] $Mod loaded." -ForegroundColor Green
}

Write-Host ""
Write-Host "  [2/5] Connecting..." -ForegroundColor Cyan
try {
    # Connect Microsoft Graph FIRST, always. ExchangeOnlineManagement bundles an older
    # Microsoft.Identity.Client (MSAL); if it loads first, Connect-MgGraph fails with
    # "Method not found: WithLogging". This happens even under PowerShell 7 (the Graph
    # SDK's assembly isolation does not cover this auth path), so Graph must initialise
    # its newer MSAL before EXO loads its copy.
    Connect-MgGraph -Scopes "User.Read.All","Group.Read.All","GroupMember.ReadWrite.All" -NoWelcome -ErrorAction Stop
    Write-Host "  [OK] Microsoft Graph connected." -ForegroundColor Green
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "  [OK] Exchange Online connected." -ForegroundColor Green
} catch {
    Write-Host "  [ERROR] Failed to connect: $_" -ForegroundColor Red
    Read-Host "`n  Press Enter to exit"; exit 1
}

# ----------------------------------------------------------
#  VERIFY USER EXISTS
# ----------------------------------------------------------
try {
    $userObj = Get-MgUser -UserId $TargetUser -Property "Id,DisplayName,UserPrincipalName,Mail" -ErrorAction Stop
    Write-Host ""
    Write-Host "  Target User : $($userObj.DisplayName) ($($userObj.UserPrincipalName))" -ForegroundColor Green
} catch {
    Write-Host "  [ERROR] User '$TargetUser' not found in Microsoft Graph." -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Read-Host "`n  Press Enter to exit"; exit 1
}

# ----------------------------------------------------------
#  DISCOVERY: MAILBOX ACCESS
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [3/5] Discovering mailbox access..." -ForegroundColor Cyan

# SendAs - single fast query. Get-RecipientPermission returns the recipient
# Name (not SMTP) in .Identity, so resolve each to its PrimarySmtpAddress.
# Entries that don't resolve are orphaned/soft-deleted recipients (their live
# ACE is already gone - they can't be removed and will clear when purged).
$sendAsEntries = @()
try {
    $sa = @(Get-RecipientPermission -Trustee $TargetUser -ResultSize Unlimited -ErrorAction Stop |
            Where-Object { $_.AccessRights -contains 'SendAs' })
    foreach ($r in $sa) {
        $rawId    = $r.Identity.ToString()
        $smtp     = $null
        $resolved = $false
        try {
            $rcpt = Get-Recipient -Identity $rawId -ErrorAction Stop
            if ($rcpt.PrimarySmtpAddress) {
                $smtp     = $rcpt.PrimarySmtpAddress.ToString()
                $resolved = $true
            }
        } catch { }
        $sendAsEntries += [PSCustomObject]@{
            RawId    = $rawId
            Smtp     = $smtp
            Resolved = $resolved
        }
    }
    Write-Host "  [OK] SendAs entries : $($sendAsEntries.Count)" -ForegroundColor Green
    $unresolvedSA = @($sendAsEntries | Where-Object { -not $_.Resolved }).Count
    if ($unresolvedSA -gt 0) {
        Write-Host "  [!]  $unresolvedSA SendAs entry(ies) point to deleted/orphaned recipients - cannot be removed." -ForegroundColor Yellow
    }
} catch {
    Write-Host "  [WARN] Could not query SendAs: $_" -ForegroundColor Yellow
}

# FullAccess - must iterate mailboxes
Write-Host "  Scanning mailboxes for FullAccess (scope: $Scope)..." -ForegroundColor DarkCyan

$mailboxFilter = if ($Scope -eq 'Shared') {
    @{ RecipientTypeDetails = 'SharedMailbox' }
} else {
    @{}
}
try {
    $allMbx = if ($Scope -eq 'Shared') {
        @(Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop)
    } else {
        @(Get-Mailbox -ResultSize Unlimited -ErrorAction Stop)
    }
} catch {
    Write-Host "  [ERROR] Could not enumerate mailboxes: $_" -ForegroundColor Red
    $allMbx = @()
}

$fullAccessMap = @{}
$total = $allMbx.Count
$i     = 0
foreach ($mbx in $allMbx) {
    $i++
    if ($total -gt 0) {
        $pct = [math]::Round(($i / $total) * 100, 1)
        Write-Progress -Activity "Scanning mailboxes for FullAccess" `
                       -Status "$i / $total ($pct%)  |  $($mbx.PrimarySmtpAddress)" `
                       -PercentComplete $pct
    }
    try {
        $perm = Get-MailboxPermission -Identity $mbx.Identity -User $TargetUser -ErrorAction SilentlyContinue |
                Where-Object { $_.AccessRights -contains 'FullAccess' -and -not $_.IsInherited -and $_.User -notlike 'NT AUTHORITY\*' }
        if ($perm) {
            $key = $mbx.PrimarySmtpAddress.ToString().ToLower()
            $fullAccessMap[$key] = $mbx
        }
    } catch { }
}
Write-Progress -Activity "Scanning mailboxes for FullAccess" -Completed
Write-Host "  [OK] FullAccess entries: $($fullAccessMap.Count)" -ForegroundColor Green

# Merge into one map keyed by mailbox SMTP (or raw name for unresolved SendAs).
# Each entry carries the exact identity to use per-right at removal time:
#   FAIdentity = mailbox SMTP (for Remove-MailboxPermission)
#   SAIdentity = original Get-RecipientPermission .Identity (for Remove-RecipientPermission)
$mbxAccess = @{}
foreach ($k in $fullAccessMap.Keys) {
    $mbx = $fullAccessMap[$k]
    $smtp = $mbx.PrimarySmtpAddress.ToString()
    $mbxAccess[$k] = [PSCustomObject]@{
        DisplayKey = $smtp
        FullAccess = $true
        SendAs     = $false
        Resolved   = $true
        FAIdentity = $smtp
        SAIdentity = $null
    }
}
foreach ($e in $sendAsEntries) {
    $key = if ($e.Resolved) { $e.Smtp.ToLower() } else { $e.RawId.ToLower() }
    if ($mbxAccess.ContainsKey($key)) {
        $mbxAccess[$key].SendAs     = $true
        $mbxAccess[$key].SAIdentity = $e.RawId
        if (-not $e.Resolved) { $mbxAccess[$key].Resolved = $false }
    } else {
        $display = if ($e.Resolved) { $e.Smtp } else { $e.RawId }
        $mbxAccess[$key] = [PSCustomObject]@{
            DisplayKey = $display
            FullAccess = $false
            SendAs     = $true
            Resolved   = $e.Resolved
            FAIdentity = $null
            SAIdentity = $e.RawId
        }
    }
}

# ----------------------------------------------------------
#  DISCOVERY: GROUP MEMBERSHIPS
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [4/5] Discovering group memberships..." -ForegroundColor Cyan

$Groups = [System.Collections.Generic.List[PSCustomObject]]::new()
try {
    $memberships = @(Get-MgUserMemberOf -UserId $userObj.Id -All -ErrorAction Stop)
    foreach ($m in $memberships) {
        $props      = $m.AdditionalProperties
        $oDataType  = $props['@odata.type']
        if ($oDataType -ne '#microsoft.graph.group') { continue }

        $displayName     = $props['displayName']
        $mail            = $props['mail']
        $mailEnabled     = [bool]$props['mailEnabled']
        $securityEnabled = [bool]$props['securityEnabled']
        $groupTypes      = @($props['groupTypes'])

        $type = if ($groupTypes -contains 'Unified')                              { 'M365 Group' }
                elseif ($mailEnabled -and $securityEnabled)                       { 'Mail-Enabled Security Group' }
                elseif ($mailEnabled -and -not $securityEnabled)                  { 'Distribution Group' }
                elseif ($securityEnabled -and -not $mailEnabled)                  { 'Security Group' }
                else                                                              { 'Unknown' }

        $Eligible = $type -in @('Distribution Group','Mail-Enabled Security Group','M365 Group')

        [void]$Groups.Add([PSCustomObject]@{
            Id          = $m.Id
            DisplayName = $displayName
            Mail        = $mail
            Type        = $type
            Eligible    = $Eligible
        })
    }
    Write-Host "  [OK] Memberships found: $($Groups.Count)  (eligible for removal: $(@($Groups | Where-Object Eligible).Count))" -ForegroundColor Green
} catch {
    Write-Host "  [WARN] Could not enumerate group memberships: $_" -ForegroundColor Yellow
}

# ----------------------------------------------------------
#  DISPLAY DISCOVERY RESULTS
# ----------------------------------------------------------
Write-Host ""
Write-Host "  +-- MAILBOX ACCESS ($($mbxAccess.Count)) ---------------------" -ForegroundColor Cyan
if ($mbxAccess.Count -eq 0) {
    Write-Host "    (none)" -ForegroundColor DarkGray
} else {
    $mbxAccess.Values | Sort-Object DisplayKey | ForEach-Object {
        $rights = @()
        if ($_.FullAccess) { $rights += 'FullAccess' }
        if ($_.SendAs)     { $rights += 'SendAs' }
        $note = if (-not $_.Resolved) { '  (orphaned/soft-deleted - cannot remove)' } else { '' }
        $col  = if (-not $_.Resolved) { 'DarkGray' } else { 'White' }
        Write-Host ("    {0,-50} [{1}]{2}" -f $_.DisplayKey, ($rights -join ' + '), $note) -ForegroundColor $col
    }
}

Write-Host ""
Write-Host "  +-- GROUP MEMBERSHIPS ($($Groups.Count)) -------------------" -ForegroundColor Cyan
if ($Groups.Count -eq 0) {
    Write-Host "    (none)" -ForegroundColor DarkGray
} else {
    $Groups | Sort-Object Type, DisplayName | ForEach-Object {
        $tag = if ($_.Eligible) { '*' } else { ' ' }
        $id  = if ($_.Mail) { $_.Mail } else { '(no SMTP)' }
        Write-Host ("   $tag {0,-25} {1,-45} {2}" -f $_.Type, $_.DisplayName, $id)
    }
    Write-Host "    (*) eligible for removal here. Pure Security Groups are excluded." -ForegroundColor DarkGray
}

# Write the audit CSV (always - even if user removes nothing)
$AuditRows = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($v in $mbxAccess.Values) {
    $detailParts = @()
    if ($v.FullAccess) { $detailParts += 'FullAccess' }
    if ($v.SendAs)     { $detailParts += 'SendAs' }
    [void]$AuditRows.Add([PSCustomObject]@{
        Category   = 'Mailbox'
        Identity   = $v.DisplayKey
        Detail     = ($detailParts -join ' + ')
        Eligible   = $v.Resolved
    })
}
foreach ($g in $Groups) {
    $gIdentity = if ($g.Mail) { $g.Mail } else { $g.DisplayName }
    [void]$AuditRows.Add([PSCustomObject]@{
        Category   = 'Group'
        Identity   = $gIdentity
        Detail     = $g.Type
        Eligible   = $g.Eligible
    })
}
try {
    $AuditRows | Export-Csv -Path $AuditCsv -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "  Audit CSV : $AuditCsv" -ForegroundColor DarkGray
} catch {
    Write-Host "  [WARN] Could not write audit CSV: $_" -ForegroundColor Yellow
}

# ----------------------------------------------------------
#  HELPER - read a multiline paste block
# ----------------------------------------------------------
function Read-PasteBlock {
    param ([string]$Prompt)
    Write-Host ""
    Write-Host "  $Prompt" -ForegroundColor Yellow
    Write-Host "  (blank line = SKIP this phase   /   type 'all' alone = target everything)" -ForegroundColor DarkGray
    Write-Host ""
    $lines = @()
    while ($true) {
        $l = Read-Host
        if ($l -eq "") { break }
        $lines += $l
    }
    return $lines
}

function Resolve-Selection {
    param (
        [string[]]$RawLines,
        [string[]]$AllIdentities      # the full discovered set, lower-case
    )
    if ($RawLines.Count -eq 0) {
        return @{ Mode = 'Skip'; Targets = @(); Unmatched = @() }
    }
    if ($RawLines.Count -eq 1 -and $RawLines[0].Trim().ToLower() -eq 'all') {
        return @{ Mode = 'All'; Targets = $AllIdentities; Unmatched = @() }
    }
    # Parse pasted entries, dedupe, intersect with discovered set
    $targets   = [System.Collections.Generic.List[string]]::new()
    $seen      = @{}
    $unmatched = [System.Collections.Generic.List[string]]::new()
    foreach ($line in $RawLines) {
        foreach ($candidate in ($line -split '[,;\s]+')) {
            $val = $candidate.Trim().ToLower()
            if (-not $val) { continue }
            if ($seen.ContainsKey($val)) { continue }
            $seen[$val] = $true
            if ($AllIdentities -contains $val) {
                [void]$targets.Add($val)
            } else {
                [void]$unmatched.Add($val)
            }
        }
    }
    return @{ Mode = 'Selected'; Targets = $targets.ToArray(); Unmatched = $unmatched.ToArray() }
}

# ----------------------------------------------------------
#  PHASE A - MAILBOX ACCESS REMOVAL
# ----------------------------------------------------------
$RemovalResults = [System.Collections.Generic.List[PSCustomObject]]::new()

$mbxKeys = @($mbxAccess.Keys | Sort-Object)
$mbxRaw  = Read-PasteBlock -Prompt "Paste MAILBOX emails to remove access from"
$mbxSel  = Resolve-Selection -RawLines $mbxRaw -AllIdentities $mbxKeys

if ($mbxSel.Mode -eq 'Skip' -or $mbxSel.Targets.Count -eq 0) {
    if ($mbxSel.Mode -ne 'Skip') {
        Write-Host "  [!] No pasted entries matched discovered mailboxes. Skipping mailbox phase." -ForegroundColor Yellow
        if ($mbxSel.Unmatched.Count -gt 0) {
            Write-Host "  Unmatched:" -ForegroundColor DarkGray
            $mbxSel.Unmatched | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
        }
    } else {
        Write-Host "  Mailbox phase skipped." -ForegroundColor DarkGray
    }
} else {
    Write-Host ""
    Write-Host "  Will remove access from $($mbxSel.Targets.Count) mailbox(es):" -ForegroundColor Cyan
    $mbxSel.Targets | ForEach-Object { Write-Host "    $_" }
    if ($mbxSel.Unmatched.Count -gt 0) {
        Write-Host "  Ignored (not in discovered list):" -ForegroundColor Yellow
        $mbxSel.Unmatched | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
    }
    $ok = Read-Host "`n  Proceed with mailbox removal? (Y/N)"
    if ($ok -in @('Y','y')) {
        foreach ($mbxKey in $mbxSel.Targets) {
            $info     = $mbxAccess[$mbxKey]
            $faOk     = -not $info.FullAccess
            $saOk     = -not $info.SendAs
            $faError  = $null
            $saError  = $null
            $notes    = @()
            $noAce    = $false   # a removal warned that the ACE wasn't present

            if ($info.FullAccess) {
                $wv = $null
                try {
                    Remove-MailboxPermission -Identity $info.FAIdentity -User $TargetUser -AccessRights FullAccess -Confirm:$false -WarningVariable wv -WarningAction SilentlyContinue -ErrorAction Stop | Out-Null
                    $faOk = $true
                    if ($wv -and (($wv -join ' ') -match "ACE doesn't exist|does not exist")) {
                        $notes += 'FullAccess ACE not present (nothing to remove)'; $noAce = $true
                    }
                } catch { $faError = $_.Exception.Message }
            }
            if ($info.SendAs) {
                $wv = $null
                try {
                    Remove-RecipientPermission -Identity $info.SAIdentity -Trustee $TargetUser -AccessRights SendAs -Confirm:$false -WarningVariable wv -WarningAction SilentlyContinue -ErrorAction Stop | Out-Null
                    $saOk = $true
                    if ($wv -and (($wv -join ' ') -match "ACE doesn't exist|does not exist")) {
                        $notes += 'SendAs ACE not present (orphaned/soft-deleted - nothing to remove)'; $noAce = $true
                    }
                } catch { $saError = $_.Exception.Message }
            }

            $status = if ($faOk -and $saOk) {
                if ($info.FullAccess -and $info.SendAs) { 'Removed-Both' }
                elseif ($info.FullAccess)               { 'Removed-FullAccess' }
                else                                    { 'Removed-SendAs' }
            } elseif (-not $faOk -and -not $saOk) { 'Failed-Both' }
              elseif (-not $faOk)                  { 'Failed-FullAccess' }
              else                                  { 'Failed-SendAs' }

            # Nothing was actually present to remove (orphaned grant) - don't claim a removal.
            if ($status -like 'Removed-*' -and $noAce) { $status = 'Skipped-NoACE' }

            $msgParts = @()
            if ($faError) { $msgParts += "FA: $faError" }
            if ($saError) { $msgParts += "SA: $saError" }
            $msgParts += $notes
            $msg = $msgParts -join ' | '

            $detailStr = if ($info.FullAccess -and $info.SendAs) { 'FullAccess + SendAs' }
                         elseif ($info.FullAccess)                { 'FullAccess' }
                         else                                     { 'SendAs' }

            $RemovalResults.Add([PSCustomObject]@{
                Category = 'Mailbox'; Identity = $info.DisplayKey; Detail = $detailStr
                Status   = $status; ErrorMessage = $msg
            })

            $col = if ($status -like 'Removed-*') { 'Green' } elseif ($status -like 'Skipped*') { 'Yellow' } else { 'Red' }
            Write-Host "  $status : $($info.DisplayKey)" -ForegroundColor $col
        }
    } else {
        Write-Host "  Mailbox phase aborted." -ForegroundColor Red
    }
}

# ----------------------------------------------------------
#  PHASE B - GROUP MEMBERSHIP REMOVAL
# ----------------------------------------------------------
$eligibleGroups = @($Groups | Where-Object Eligible)
# Identity key for groups: prefer Mail, fall back to DisplayName
$groupKeys      = @($eligibleGroups | ForEach-Object {
    $idForKey = if ($_.Mail) { $_.Mail } else { $_.DisplayName }
    $idForKey.ToLower()
})
$groupLookup    = @{}
foreach ($g in $eligibleGroups) {
    $idForKey = if ($g.Mail) { $g.Mail } else { $g.DisplayName }
    $k = $idForKey.ToLower()
    $groupLookup[$k] = $g
}

$grpRaw = Read-PasteBlock -Prompt "Paste GROUP SMTP addresses (or names) to remove membership from"
$grpSel = Resolve-Selection -RawLines $grpRaw -AllIdentities $groupKeys

if ($grpSel.Mode -eq 'Skip' -or $grpSel.Targets.Count -eq 0) {
    if ($grpSel.Mode -ne 'Skip') {
        Write-Host "  [!] No pasted entries matched eligible groups. Skipping group phase." -ForegroundColor Yellow
        if ($grpSel.Unmatched.Count -gt 0) {
            Write-Host "  Unmatched:" -ForegroundColor DarkGray
            $grpSel.Unmatched | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
        }
    } else {
        Write-Host "  Group phase skipped." -ForegroundColor DarkGray
    }
} else {
    Write-Host ""
    Write-Host "  Will remove user from $($grpSel.Targets.Count) group(s):" -ForegroundColor Cyan
    foreach ($k in $grpSel.Targets) {
        $g = $groupLookup[$k]
        Write-Host ("    {0,-25} {1}" -f $g.Type, $g.DisplayName)
    }
    if ($grpSel.Unmatched.Count -gt 0) {
        Write-Host "  Ignored (not in eligible group list):" -ForegroundColor Yellow
        $grpSel.Unmatched | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
    }
    $ok = Read-Host "`n  Proceed with group removal? (Y/N)"
    if ($ok -in @('Y','y')) {
        foreach ($k in $grpSel.Targets) {
            $g = $groupLookup[$k]
            $idForCmd = if ($g.Mail) { $g.Mail } else { $g.Id }

            try {
                switch ($g.Type) {
                    'Distribution Group'           { Remove-DistributionGroupMember -Identity $idForCmd -Member $TargetUser -Confirm:$false -ErrorAction Stop }
                    'Mail-Enabled Security Group'  { Remove-DistributionGroupMember -Identity $idForCmd -Member $TargetUser -Confirm:$false -ErrorAction Stop }
                    'M365 Group'                   { Remove-UnifiedGroupLinks       -Identity $idForCmd -LinkType Members -Links $TargetUser -Confirm:$false -ErrorAction Stop }
                    default                        { throw "Unsupported group type '$($g.Type)' for removal." }
                }
                $RemovalResults.Add([PSCustomObject]@{
                    Category = 'Group'; Identity = $idForCmd; Detail = $g.Type
                    Status   = 'Removed'; ErrorMessage = $null
                })
                Write-Host "  Removed : $($g.Type) - $($g.DisplayName)" -ForegroundColor Green
            } catch {
                $RemovalResults.Add([PSCustomObject]@{
                    Category = 'Group'; Identity = $idForCmd; Detail = $g.Type
                    Status   = 'Failed'; ErrorMessage = $_.Exception.Message
                })
                Write-Host "  Failed  : $($g.DisplayName) - $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    } else {
        Write-Host "  Group phase aborted." -ForegroundColor Red
    }
}

# ----------------------------------------------------------
#  EXPORT REMOVAL LOG + DISCONNECT
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [5/5] Exporting removal log..." -ForegroundColor Cyan
try {
    if ($RemovalResults.Count -gt 0) {
        $RemovalResults | Export-Csv -Path $RemovalCsv -NoTypeInformation -Encoding UTF8
        Write-Host "  [OK] Removal CSV : $RemovalCsv" -ForegroundColor Green
    } else {
        Write-Host "  No removal actions executed - skipping removal CSV." -ForegroundColor DarkGray
    }
} catch {
    Write-Host "  [ERROR] Failed to write removal CSV: $_" -ForegroundColor Red
}

Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Disconnect-MgGraph -ErrorAction SilentlyContinue
Write-Host "  Sessions disconnected." -ForegroundColor DarkGray

# ----------------------------------------------------------
#  SUMMARY
# ----------------------------------------------------------
$RemovedCnt  = ($RemovalResults | Where-Object { $_.Status -like 'Removed*' }).Count
$SkippedCnt  = ($RemovalResults | Where-Object { $_.Status -like 'Skipped*' }).Count
$FailedCnt   = ($RemovalResults | Where-Object { $_.Status -like 'Failed*'  }).Count

Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |                  RUN SUMMARY                    |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Tenant Code         : {0,-27}|" -f $TenantCode) -ForegroundColor White
Write-Host ("  | Target User         : {0,-27}|" -f $TargetUser) -ForegroundColor White
Write-Host ("  | Mailbox Access Found: {0,-27}|" -f $mbxAccess.Count) -ForegroundColor White
Write-Host ("  | Group Memberships   : {0,-27}|" -f $Groups.Count) -ForegroundColor White
Write-Host ("  |   (eligible to remove): {0,-25}|" -f (@($Groups | Where-Object Eligible).Count)) -ForegroundColor White
Write-Host ("  | Actions Executed    : {0,-27}|" -f $RemovalResults.Count) -ForegroundColor White
Write-Host ("  |   Removed           : {0,-27}|" -f $RemovedCnt)  -ForegroundColor White
Write-Host ("  |   Skipped (no ACE)  : {0,-27}|" -f $SkippedCnt)  -ForegroundColor White
Write-Host ("  |   Failed            : {0,-27}|" -f $FailedCnt)   -ForegroundColor White
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""
Read-Host "  Press Enter to exit"
