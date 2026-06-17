#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Reports delegated calendar permissions across all mailboxes in a tenant,
    classifying each calendar's mailbox type.

.DESCRIPTION
    Connects to Exchange Online, enumerates every mailbox, and reads the
    calendar folder permissions. Default and Anonymous entries are excluded.
    Each row records the mailbox type (User / Shared / Room / Equipment), the
    mailbox, the delegate, and the access rights.

    The calendar folder is read as ':\Calendar' first (fast); if that fails the
    real (possibly localized) calendar folder name is resolved and retried.

    Results are exported to C:\MSP-M365-Utility\.

.NOTES
    Required Modules:
        - ExchangeOnlineManagement
    Required Permissions:
        - Exchange Online : View-Only Recipients (or higher)
#>

# ----------------------------------------------------------
#  CONFIGURATION
# ----------------------------------------------------------
$ErrorActionPreference = 'Stop'
$Host.UI.RawUI.WindowTitle = "M365 Calendar Access Report"

# ----------------------------------------------------------
#  BANNER
# ----------------------------------------------------------
Clear-Host
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |        M365 Calendar Access Report Script        |" -ForegroundColor Cyan
Write-Host "  |              Exchange Online                     |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""

# ----------------------------------------------------------
#  INPUT - THREE-LETTER TENANT CODE
# ----------------------------------------------------------
do {
    $TenantCode = (Read-Host "  Enter the three-letter Tenant Code (e.g. ABC)").Trim().ToUpper()
    if ($TenantCode -notmatch '^[A-Z]{3}$') {
        Write-Host "  [!] Invalid input. Please enter exactly 3 letters (A-Z)." -ForegroundColor Yellow
    }
} while ($TenantCode -notmatch '^[A-Z]{3}$')

# ----------------------------------------------------------
#  OUTPUT PATH (C:\MSP-M365-Utility\)
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
$OutputFile = "CalendarAccess_${TenantCode}_${Timestamp}.csv"
$OutputPath = Join-Path $OutputRoot $OutputFile

Write-Host ""
Write-Host "  Tenant Code : $TenantCode" -ForegroundColor Green
Write-Host "  Output File : $OutputPath" -ForegroundColor Green
Write-Host ""

# ----------------------------------------------------------
#  TIMER
# ----------------------------------------------------------
$ScriptStart = [System.Diagnostics.Stopwatch]::StartNew()

# ----------------------------------------------------------
#  MODULE CHECK
# ----------------------------------------------------------
Write-Host "  [1/4] Checking required modules..." -ForegroundColor Cyan
foreach ($Mod in @('ExchangeOnlineManagement')) {
    if (-not (Get-Module -ListAvailable -Name $Mod)) {
        Write-Host "  [!] Module '$Mod' not found. Installing..." -ForegroundColor Yellow
        Install-Module -Name $Mod -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module -Name $Mod -ErrorAction Stop
    Write-Host "  [OK] $Mod loaded." -ForegroundColor Green
}

# ----------------------------------------------------------
#  CONNECT
# ----------------------------------------------------------
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
#  RETRIEVE MAILBOXES
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [3/4] Retrieving mailboxes..." -ForegroundColor Cyan
try {
    $Mailboxes  = @(Get-Mailbox -ResultSize Unlimited -ErrorAction Stop)
    $TotalCount = $Mailboxes.Count
    Write-Host "  [OK] Found $TotalCount mailbox(es)." -ForegroundColor Green
} catch {
    Write-Host "  [ERROR] Could not retrieve mailboxes: $_" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Read-Host "`n  Press Enter to exit"; exit 1
}

# ----------------------------------------------------------
#  HELPER - friendly mailbox type
# ----------------------------------------------------------
function Get-MailboxTypeLabel {
    param([string]$RecipientTypeDetails)
    switch ($RecipientTypeDetails) {
        'UserMailbox'      { 'User'      ; break }
        'SharedMailbox'    { 'Shared'    ; break }
        'RoomMailbox'      { 'Room'      ; break }
        'EquipmentMailbox' { 'Equipment' ; break }
        default            { $RecipientTypeDetails }
    }
}

# ----------------------------------------------------------
#  PROCESS WITH PROGRESS BAR
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [4/4] Reading calendar permissions..." -ForegroundColor Cyan
Write-Host ""

$Results            = [System.Collections.Generic.List[PSCustomObject]]::new()
$Counter            = 0
$MailboxesWithAccess = 0

foreach ($Mailbox in $Mailboxes) {
    $Counter++
    $PercentComplete = [math]::Round(($Counter / $TotalCount) * 100, 1)
    $BarWidth        = 40
    $FilledBars      = [math]::Floor($PercentComplete / 100 * $BarWidth)
    $EmptyBars       = $BarWidth - $FilledBars
    $ProgressBar     = "[" + ("#" * $FilledBars) + ("-" * $EmptyBars) + "]"

    $Smtp = $Mailbox.PrimarySmtpAddress.ToString()
    $Type = Get-MailboxTypeLabel $Mailbox.RecipientTypeDetails

    Write-Progress `
        -Activity        "Calendar Access -- Tenant: $TenantCode" `
        -Status          "$ProgressBar  $Counter / $TotalCount  ($PercentComplete%)  |  $($Mailbox.DisplayName)" `
        -PercentComplete  $PercentComplete `
        -CurrentOperation "Type: $Type  |  $Smtp"

    # Read calendar permissions - try :\Calendar, then resolve localized folder name
    $Permissions = $null
    try {
        $Permissions = @(Get-MailboxFolderPermission "${Smtp}:\Calendar" -ErrorAction Stop)
    } catch {
        try {
            $calFolder = Get-MailboxFolderStatistics -Identity $Smtp -FolderScope Calendar -ErrorAction Stop |
                Where-Object { $_.FolderType -eq 'Calendar' } | Select-Object -First 1
            if ($calFolder) {
                $Permissions = @(Get-MailboxFolderPermission "${Smtp}:\$($calFolder.Name)" -ErrorAction Stop)
            }
        } catch {
            Write-Host "  [WARN] Could not read calendar for $Smtp : $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    if (-not $Permissions) { continue }

    $hadDelegate = $false
    foreach ($Permission in $Permissions) {
        $u = "$($Permission.User)"
        if ($u -like '*Anonymous*' -or $u -like '*Default*') { continue }
        if (($Permission.AccessRights -join '') -match '^None$') { continue }

        $hadDelegate = $true
        $Results.Add([PSCustomObject]@{
            MailboxType  = $Type
            Mailbox      = $Mailbox.DisplayName
            EmailAddress = $Smtp
            User         = $u
            AccessRights = ($Permission.AccessRights -join ', ')
        })
    }
    if ($hadDelegate) { $MailboxesWithAccess++ }
}

Write-Progress -Activity "Calendar Access" -Completed

# ----------------------------------------------------------
#  EXPORT
# ----------------------------------------------------------
try {
    $Results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "  [OK] CSV exported successfully." -ForegroundColor Green
    Write-Host "       Path: $OutputPath"          -ForegroundColor DarkGray
} catch {
    Write-Host "  [ERROR] Failed to export CSV: $_" -ForegroundColor Red
}

# ----------------------------------------------------------
#  DISCONNECT
# ----------------------------------------------------------
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Write-Host "  Session disconnected." -ForegroundColor DarkGray

# ----------------------------------------------------------
#  SUMMARY
# ----------------------------------------------------------
$ScriptStart.Stop()
$Elapsed = $ScriptStart.Elapsed

$UserRows      = ($Results | Where-Object { $_.MailboxType -eq 'User'      }).Count
$SharedRows    = ($Results | Where-Object { $_.MailboxType -eq 'Shared'    }).Count
$RoomRows      = ($Results | Where-Object { $_.MailboxType -eq 'Room'      }).Count
$EquipmentRows = ($Results | Where-Object { $_.MailboxType -eq 'Equipment' }).Count

$RunTime = "{0:D2}h {1:D2}m {2:D2}s {3:D3}ms" -f `
    $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds, $Elapsed.Milliseconds

Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |                  RUN SUMMARY                    |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Tenant Code             : {0,-23}|" -f $TenantCode)          -ForegroundColor White
Write-Host ("  | Mailboxes Scanned       : {0,-23}|" -f $TotalCount)          -ForegroundColor White
Write-Host ("  | Mailboxes w/ Delegates  : {0,-23}|" -f $MailboxesWithAccess) -ForegroundColor White
Write-Host ("  | Access Entries (rows)   : {0,-23}|" -f $Results.Count)       -ForegroundColor White
Write-Host ("  |   on User calendars     : {0,-23}|" -f $UserRows)            -ForegroundColor White
Write-Host ("  |   on Shared calendars   : {0,-23}|" -f $SharedRows)          -ForegroundColor White
Write-Host ("  |   on Room calendars     : {0,-23}|" -f $RoomRows)            -ForegroundColor White
Write-Host ("  |   on Equipment calendars: {0,-23}|" -f $EquipmentRows)       -ForegroundColor White
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Total Run Time          : {0,-23}|" -f $RunTime)             -ForegroundColor Yellow
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""
Read-Host "  Press Enter to exit"
