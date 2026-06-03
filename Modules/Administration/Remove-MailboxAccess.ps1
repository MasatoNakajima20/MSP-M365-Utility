#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Bulk-remove FullAccess and SendAs from a target mailbox for one or more
    users.

.DESCRIPTION
    Connects to Exchange Online and, for each user pasted in, removes both
    FullAccess and SendAs on the target mailbox.

    For each user:
        - Both rights present       -> Removed-Both
        - Only one present          -> Removed-FullAccessOnly / Removed-SendAsOnly
        - Neither present           -> Not-Found
        - One or both fail to remove -> Failed-FullAccess / Failed-SendAs / Failed-Both

    Results are written to C:\MSP-M365-Utility\.

.NOTES
    Required Modules:
        - ExchangeOnlineManagement
    Required Permissions:
        - Exchange Online : Recipient Management (or higher)
#>

# ----------------------------------------------------------
#  CONFIGURATION
# ----------------------------------------------------------
$ErrorActionPreference = 'Stop'
$Host.UI.RawUI.WindowTitle = "M365 Bulk Remove Mailbox Access"

# ----------------------------------------------------------
#  BANNER
# ----------------------------------------------------------
Clear-Host
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |    M365 Bulk Remove Mailbox Access Script        |" -ForegroundColor Cyan
Write-Host "  |   FullAccess + SendAs  *  Exchange Online        |" -ForegroundColor Cyan
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
#  INPUT - TARGET MAILBOX
# ----------------------------------------------------------
$EmailRegex = '^[^@\s]+@[^@\s]+\.[^@\s]+$'
do {
    $TargetMailbox = (Read-Host "  Enter the target MAILBOX email").Trim()
    if ($TargetMailbox -notmatch $EmailRegex) {
        Write-Host "  [!] That doesn't look like an email address." -ForegroundColor Yellow
    }
} while ($TargetMailbox -notmatch $EmailRegex)

# ----------------------------------------------------------
#  OUTPUT PATH
# ----------------------------------------------------------
$OutputRoot = 'C:\MSP-M365-Utility'
if (-not (Test-Path $OutputRoot)) {
    try {
        New-Item -ItemType Directory -Path $OutputRoot -Force -ErrorAction Stop | Out-Null
    } catch {
        Write-Host "  [ERROR] Could not create output folder '$OutputRoot': $_" -ForegroundColor Red
        Read-Host "`n  Press Enter to exit"; exit 1
    }
}
$Timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputFile = "MailboxAccessRemove_${TenantCode}_${Timestamp}.csv"
$OutputPath = Join-Path $OutputRoot $OutputFile

# ----------------------------------------------------------
#  PASTE-LIST INPUT
# ----------------------------------------------------------
Write-Host ""
Write-Host "  Paste user emails below (one per line)." -ForegroundColor Yellow
Write-Host "  Each user will have BOTH FullAccess and SendAs removed from $TargetMailbox." -ForegroundColor DarkGray
Write-Host "  When finished, press ENTER on a blank line." -ForegroundColor Yellow
Write-Host ""

$RawLines = @()
while ($true) {
    $line = Read-Host
    if ($line -eq "") { break }
    $RawLines += $line
}

# ----------------------------------------------------------
#  PARSE & VALIDATE
# ----------------------------------------------------------
$Users      = [System.Collections.Generic.List[string]]::new()
$Malformed  = [System.Collections.Generic.List[string]]::new()
$SeenEmails = @{}

foreach ($line in $RawLines) {
    $trimmed = $line.Trim()
    if (-not $trimmed) { continue }

    foreach ($candidate in ($trimmed -split '[,;\s]+')) {
        $email = $candidate.Trim()
        if (-not $email) { continue }
        if ($email -notmatch $EmailRegex) {
            [void]$Malformed.Add($email); continue
        }
        $key = $email.ToLower()
        if ($SeenEmails.ContainsKey($key)) { continue }
        $SeenEmails[$key] = $true
        [void]$Users.Add($email)
    }
}

if ($Users.Count -eq 0) {
    Write-Host ""
    Write-Host "  [!] No valid emails found. Nothing to do." -ForegroundColor Yellow
    if ($Malformed.Count -gt 0) {
        Write-Host "  Malformed entries:" -ForegroundColor DarkGray
        $Malformed | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
    }
    Read-Host "`n  Press Enter to exit"; exit 0
}

# ----------------------------------------------------------
#  MODULE CHECK + CONNECT
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [1/4] Checking required modules..." -ForegroundColor Cyan
foreach ($Mod in @('ExchangeOnlineManagement')) {
    if (-not (Get-Module -ListAvailable -Name $Mod)) {
        Write-Host "  [!] Module '$Mod' not found. Installing..." -ForegroundColor Yellow
        Install-Module -Name $Mod -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module -Name $Mod -ErrorAction Stop
    Write-Host "  [OK] $Mod loaded." -ForegroundColor Green
}

Write-Host ""
Write-Host "  [2/4] Connecting to Exchange Online..." -ForegroundColor Cyan
try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "  [OK] Exchange Online connected." -ForegroundColor Green
} catch {
    Write-Host "  [ERROR] Failed to connect: $_" -ForegroundColor Red
    Read-Host "`n  Press Enter to exit"; exit 1
}

# ----------------------------------------------------------
#  CONFIRM TARGET EXISTS
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [3/4] Checking target mailbox..." -ForegroundColor Cyan
try {
    $existingMbx = Get-Mailbox -Identity $TargetMailbox -ErrorAction Stop
    Write-Host "  [OK] Found: $($existingMbx.DisplayName) ($($existingMbx.RecipientTypeDetails))" -ForegroundColor Green
} catch {
    Write-Host "  [ERROR] Mailbox '$TargetMailbox' not found - cannot remove permissions." -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Read-Host "`n  Press Enter to exit"; exit 1
}

# ----------------------------------------------------------
#  PREVIEW
# ----------------------------------------------------------
Write-Host ""
Write-Host "  Tenant Code      : $TenantCode" -ForegroundColor Cyan
Write-Host "  Target Mailbox   : $TargetMailbox" -ForegroundColor Cyan
Write-Host "  Users to REMOVE  : $($Users.Count)" -ForegroundColor Cyan
Write-Host "  Malformed Lines  : $($Malformed.Count)" -ForegroundColor $(if ($Malformed.Count) { 'Yellow' } else { 'DarkGray' })
Write-Host ""
$Users | ForEach-Object { Write-Host "    $_" }
if ($Malformed.Count -gt 0) {
    Write-Host ""
    Write-Host "  Skipped (malformed):" -ForegroundColor Yellow
    $Malformed | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
}

$Confirm = Read-Host "`n  Proceed with REMOVAL of FullAccess and SendAs? (Y/N)"
if ($Confirm -notin @('Y','y')) {
    Write-Host "  Aborted." -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Read-Host "`n  Press Enter to exit"; exit 0
}

$ScriptStart = [System.Diagnostics.Stopwatch]::StartNew()

# ----------------------------------------------------------
#  PROCESS
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [4/4] Removing permissions..." -ForegroundColor Cyan
Write-Host ""

$Results = [System.Collections.Generic.List[PSCustomObject]]::new()
$Total   = $Users.Count
$Counter = 0

foreach ($User in $Users) {
    $Counter++
    $PercentComplete = [math]::Round(($Counter / $Total) * 100, 1)
    $BarWidth        = 40
    $FilledBars      = [math]::Floor($PercentComplete / 100 * $BarWidth)
    $EmptyBars       = $BarWidth - $FilledBars
    $ProgressBar     = "[" + ("#" * $FilledBars) + ("-" * $EmptyBars) + "]"

    Write-Progress `
        -Activity        "Remove Mailbox Access -- Tenant: $TenantCode" `
        -Status          "$ProgressBar  $Counter / $Total  ($PercentComplete%)  |  $User" `
        -PercentComplete  $PercentComplete `
        -CurrentOperation "Mailbox: $TargetMailbox"

    # Probe existing rights
    $hasFA = $false
    try {
        $faList = Get-MailboxPermission -Identity $TargetMailbox -User $User -ErrorAction Stop |
            Where-Object { $_.AccessRights -contains 'FullAccess' -and -not $_.IsInherited -and $_.User -notlike 'NT AUTHORITY\*' }
        $hasFA = [bool]$faList
    } catch { $hasFA = $false }

    $hasSA = $false
    try {
        $saList = Get-RecipientPermission -Identity $TargetMailbox -Trustee $User -ErrorAction Stop |
            Where-Object { $_.AccessRights -contains 'SendAs' }
        $hasSA = [bool]$saList
    } catch { $hasSA = $false }

    if (-not $hasFA -and -not $hasSA) {
        $Results.Add([PSCustomObject]@{
            UserEmail = $User; RemovedFullAccess = $false; RemovedSendAs = $false
            Status = 'Not-Found'; ErrorMessage = $null
        })
        Write-Host "  Not-Found : $User" -ForegroundColor DarkGray
        continue
    }

    $faRemoved = $false
    $saRemoved = $false
    $faError   = $null
    $saError   = $null

    if ($hasFA) {
        try {
            Remove-MailboxPermission -Identity $TargetMailbox -User $User -AccessRights FullAccess -Confirm:$false -ErrorAction Stop | Out-Null
            $faRemoved = $true
        } catch { $faError = $_.Exception.Message }
    }

    if ($hasSA) {
        try {
            Remove-RecipientPermission -Identity $TargetMailbox -Trustee $User -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
            $saRemoved = $true
        } catch { $saError = $_.Exception.Message }
    }

    # Derive status
    $status = $null
    $msg    = $null
    $faAttempted = $hasFA
    $saAttempted = $hasSA
    $faOk = -not $faAttempted -or $faRemoved
    $saOk = -not $saAttempted -or $saRemoved

    if ($faOk -and $saOk) {
        if     ($faAttempted -and $saAttempted) { $status = 'Removed-Both' }
        elseif ($faAttempted)                   { $status = 'Removed-FullAccessOnly' }
        elseif ($saAttempted)                   { $status = 'Removed-SendAsOnly' }
    }
    elseif (-not $faOk -and -not $saOk) { $status = 'Failed-Both';      $msg = "FA: $faError | SA: $saError" }
    elseif ($faOk -and -not $saOk)      { $status = 'Failed-SendAs';    $msg = $saError }
    elseif (-not $faOk -and $saOk)      { $status = 'Failed-FullAccess';$msg = $faError }

    $Results.Add([PSCustomObject]@{
        UserEmail         = $User
        RemovedFullAccess = $faRemoved
        RemovedSendAs     = $saRemoved
        Status            = $status
        ErrorMessage      = $msg
    })

    $colour = if ($status -like 'Removed-*') { 'Green' } elseif ($status -like 'Failed-*') { 'Red' } else { 'Yellow' }
    Write-Host "  $status : $User" -ForegroundColor $colour
}

Write-Progress -Activity "Remove Mailbox Access" -Completed

# ----------------------------------------------------------
#  EXPORT RESULTS
# ----------------------------------------------------------
try {
    $Results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "  [OK] Results CSV exported." -ForegroundColor Green
    Write-Host "       Path: $OutputPath"     -ForegroundColor DarkGray
} catch {
    Write-Host "  [ERROR] Failed to export results: $_" -ForegroundColor Red
}

Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Write-Host "  Session disconnected." -ForegroundColor DarkGray

# ----------------------------------------------------------
#  SUMMARY
# ----------------------------------------------------------
$ScriptStart.Stop()
$Elapsed = $ScriptStart.Elapsed

$BothCnt     = ($Results | Where-Object { $_.Status -eq 'Removed-Both'           }).Count
$FAOnlyCnt   = ($Results | Where-Object { $_.Status -eq 'Removed-FullAccessOnly' }).Count
$SAOnlyCnt   = ($Results | Where-Object { $_.Status -eq 'Removed-SendAsOnly'     }).Count
$NotFoundCnt = ($Results | Where-Object { $_.Status -eq 'Not-Found'              }).Count
$FailedCnt   = ($Results | Where-Object { $_.Status -like 'Failed-*'             }).Count

$RunTime = "{0:D2}h {1:D2}m {2:D2}s {3:D3}ms" -f `
    $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds, $Elapsed.Milliseconds

Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |                  RUN SUMMARY                    |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Tenant Code         : {0,-27}|" -f $TenantCode)    -ForegroundColor White
Write-Host ("  | Target Mailbox      : {0,-27}|" -f $TargetMailbox) -ForegroundColor White
Write-Host ("  | Total Users         : {0,-27}|" -f $Total)         -ForegroundColor White
Write-Host ("  |   Removed Both      : {0,-27}|" -f $BothCnt)       -ForegroundColor White
Write-Host ("  |   Removed FA only   : {0,-27}|" -f $FAOnlyCnt)     -ForegroundColor White
Write-Host ("  |   Removed SA only   : {0,-27}|" -f $SAOnlyCnt)     -ForegroundColor White
Write-Host ("  |   Not-Found         : {0,-27}|" -f $NotFoundCnt)   -ForegroundColor White
Write-Host ("  |   Failed (any)      : {0,-27}|" -f $FailedCnt)     -ForegroundColor White
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Total Run Time      : {0,-27}|" -f $RunTime)       -ForegroundColor Yellow
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""
Read-Host "  Press Enter to exit"
