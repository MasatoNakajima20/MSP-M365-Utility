#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Remove calendar permissions on a target mailbox from one or more users.

.DESCRIPTION
    Connects to Exchange Online and removes folder permissions on the target
    mailbox's Calendar folder for each user in the pasted list.

    Paste format (one entry per line):

        user@domain.com

    Behaviour:
        - The Calendar folder name is resolved per-mailbox so this works in
          non-English tenants too.
        - One overall Y/N confirmation at the preview stage; after that each
          permission is removed straight away (Remove-MailboxFolderPermission
          -Confirm:$false).
        - If the user has no existing permission on the calendar, the row is
          logged as 'Not-Found' rather than failing.
        - Results are written to C:\MSP-M365-Utility\.

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
$Host.UI.RawUI.WindowTitle = "M365 Remove Calendar Permission"

# ----------------------------------------------------------
#  BANNER
# ----------------------------------------------------------
Clear-Host
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |      M365 Remove Calendar Permission Script      |" -ForegroundColor Cyan
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
#  INPUT - TARGET MAILBOX
# ----------------------------------------------------------
$EmailRegex = '^[^@\s]+@[^@\s]+\.[^@\s]+$'
do {
    $TargetMailbox = (Read-Host "  Enter the target MAILBOX (whose calendar to modify)").Trim()
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
$OutputFile = "CalendarPermRemove_${TenantCode}_${Timestamp}.csv"
$OutputPath = Join-Path $OutputRoot $OutputFile

# ----------------------------------------------------------
#  PASTE-LIST INPUT
# ----------------------------------------------------------
Write-Host ""
Write-Host "  Paste user emails below (one per line)." -ForegroundColor Yellow
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

    # Allow same loose splitting as Remove-DistroMember (commas / semicolons / whitespace)
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
#  PREVIEW
# ----------------------------------------------------------
Write-Host ""
Write-Host "  Target Mailbox  : $TargetMailbox" -ForegroundColor Cyan
Write-Host "  Users to REMOVE : $($Users.Count)" -ForegroundColor Cyan
Write-Host "  Malformed       : $($Malformed.Count)" -ForegroundColor $(if ($Malformed.Count) { 'Yellow' } else { 'DarkGray' })
Write-Host ""
$Users | ForEach-Object { Write-Host "    $_" }
if ($Malformed.Count -gt 0) {
    Write-Host ""
    Write-Host "  Skipped (malformed):" -ForegroundColor Yellow
    $Malformed | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
}

$Confirm = Read-Host "`n  Proceed with REMOVAL? (Y/N)"
if ($Confirm -notin @('Y','y')) {
    Write-Host "  Aborted." -ForegroundColor Red
    Read-Host "`n  Press Enter to exit"; exit 0
}

$ScriptStart = [System.Diagnostics.Stopwatch]::StartNew()

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
#  RESOLVE CALENDAR FOLDER PATH
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [3/4] Resolving Calendar folder for $TargetMailbox ..." -ForegroundColor Cyan
try {
    $CalendarFolder = Get-MailboxFolderStatistics -Identity $TargetMailbox -FolderScope Calendar -ErrorAction Stop |
        Where-Object { $_.FolderType -eq 'Calendar' } | Select-Object -First 1
    if (-not $CalendarFolder) {
        throw "No Calendar folder returned for '$TargetMailbox'."
    }
    $CalendarPath = "${TargetMailbox}:\$($CalendarFolder.Name)"
    Write-Host "  [OK] Calendar folder : $CalendarPath" -ForegroundColor Green
} catch {
    Write-Host "  [ERROR] Could not resolve Calendar folder: $_" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Read-Host "`n  Press Enter to exit"; exit 1
}

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
        -Activity        "Remove Calendar Permission -- Tenant: $TenantCode" `
        -Status          "$ProgressBar  $Counter / $Total  ($PercentComplete%)  |  $User" `
        -PercentComplete  $PercentComplete `
        -CurrentOperation "Mailbox: $TargetMailbox"

    # Check first - so we can report 'Not-Found' cleanly
    $existing = $null
    try {
        $existing = Get-MailboxFolderPermission -Identity $CalendarPath -User $User -ErrorAction Stop
    } catch {
        $existing = $null
    }

    if (-not $existing) {
        $Results.Add([PSCustomObject]@{
            UserEmail        = $User
            PreviousAccess   = $null
            Status           = 'Not-Found'
            ErrorMessage     = $null
        })
        Write-Host "  Not-Found : $User" -ForegroundColor DarkGray
        continue
    }

    $previousAccess = ($existing.AccessRights -join ',')

    try {
        Remove-MailboxFolderPermission -Identity $CalendarPath -User $User -Confirm:$false -ErrorAction Stop
        $Results.Add([PSCustomObject]@{
            UserEmail        = $User
            PreviousAccess   = $previousAccess
            Status           = 'Removed'
            ErrorMessage     = $null
        })
        Write-Host "  Removed   : $User  (was: $previousAccess)" -ForegroundColor Green
    } catch {
        $Results.Add([PSCustomObject]@{
            UserEmail        = $User
            PreviousAccess   = $previousAccess
            Status           = 'Failed'
            ErrorMessage     = $_.Exception.Message
        })
        Write-Host "  Failed    : $User - $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Progress -Activity "Remove Calendar Permission" -Completed

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

$RemovedCnt  = ($Results | Where-Object { $_.Status -eq 'Removed'   }).Count
$NotFoundCnt = ($Results | Where-Object { $_.Status -eq 'Not-Found' }).Count
$FailedCnt   = ($Results | Where-Object { $_.Status -eq 'Failed'    }).Count

$RunTime = "{0:D2}h {1:D2}m {2:D2}s {3:D3}ms" -f `
    $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds, $Elapsed.Milliseconds

Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |                  RUN SUMMARY                    |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Tenant Code      : {0,-30}|" -f $TenantCode)   -ForegroundColor White
Write-Host ("  | Target Mailbox   : {0,-30}|" -f $TargetMailbox)-ForegroundColor White
Write-Host ("  | Total Users      : {0,-30}|" -f $Total)        -ForegroundColor White
Write-Host ("  |   Removed        : {0,-30}|" -f $RemovedCnt)   -ForegroundColor White
Write-Host ("  |   Not-Found      : {0,-30}|" -f $NotFoundCnt)  -ForegroundColor White
Write-Host ("  |   Failed         : {0,-30}|" -f $FailedCnt)    -ForegroundColor White
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Total Run Time   : {0,-30}|" -f $RunTime)      -ForegroundColor Yellow
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""
Read-Host "  Press Enter to exit"
