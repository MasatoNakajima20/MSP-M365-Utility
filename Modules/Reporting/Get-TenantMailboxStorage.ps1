#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Reports mailbox storage usage across a tenant, including archive usage.

.DESCRIPTION
    Connects to Exchange Online and, for every mailbox, reports:
        - Mailbox type (User / Shared / Room / Equipment)
        - Current primary mailbox usage (GB) and item count
        - Mailbox quota cap (ProhibitSendReceiveQuota) and % used
        - Whether an online archive is enabled
        - Archive usage (GB), item count, and archive quota (GB)

    All sizes are normalised to GB from Exchange's "x.x GB (n bytes)" strings.

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
$Host.UI.RawUI.WindowTitle = "M365 Mailbox Storage Report"

# ----------------------------------------------------------
#  BANNER
# ----------------------------------------------------------
Clear-Host
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |        M365 Mailbox Storage Report Script        |" -ForegroundColor Cyan
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
$OutputFile = "MailboxStorage_${TenantCode}_${Timestamp}.csv"
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
    $Mailboxes  = @(Get-Mailbox -ResultSize Unlimited -ErrorAction Stop | Sort-Object DisplayName)
    $TotalCount = $Mailboxes.Count
    Write-Host "  [OK] Found $TotalCount mailbox(es)." -ForegroundColor Green
} catch {
    Write-Host "  [ERROR] Could not retrieve mailboxes: $_" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Read-Host "`n  Press Enter to exit"; exit 1
}

# ----------------------------------------------------------
#  HELPERS
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

# Normalise an Exchange size value ("1.5 GB (1,610,612,736 bytes)" / "Unlimited") to GB.
# Returns $null for Unlimited or unparseable values.
function ConvertTo-GB {
    param($Size)
    if ($null -eq $Size) { return $null }
    $s = $Size.ToString()
    if ($s -match 'Unlimited') { return $null }
    if ($s -match '\(([\d,]+)\s*bytes\)') {
        return [math]::Round((([double]($Matches[1] -replace ',', '')) / 1GB), 2)
    }
    if ($s -match '([\d\.]+)\s*(B|KB|MB|GB|TB)') {
        $n = [double]$Matches[1]
        switch ($Matches[2]) {
            'B'  { return [math]::Round($n / 1GB, 4) }
            'KB' { return [math]::Round($n / 1MB, 4) }
            'MB' { return [math]::Round($n / 1024, 3) }
            'GB' { return [math]::Round($n, 2) }
            'TB' { return [math]::Round($n * 1024, 2) }
        }
    }
    return $null
}

$EmptyGuid = '00000000-0000-0000-0000-000000000000'

# ----------------------------------------------------------
#  PROCESS WITH PROGRESS BAR
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [4/4] Gathering storage statistics..." -ForegroundColor Cyan
Write-Host ""

$Results = [System.Collections.Generic.List[PSCustomObject]]::new()
$Counter = 0

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
        -Activity        "Mailbox Storage -- Tenant: $TenantCode" `
        -Status          "$ProgressBar  $Counter / $TotalCount  ($PercentComplete%)  |  $($Mailbox.DisplayName)" `
        -PercentComplete  $PercentComplete `
        -CurrentOperation "Type: $Type  |  $Smtp"

    # Primary mailbox stats
    $UsageGB = $null
    $Items   = $null
    try {
        $stats   = Get-MailboxStatistics -Identity $Smtp -ErrorAction Stop
        $UsageGB = ConvertTo-GB $stats.TotalItemSize
        $Items   = $stats.ItemCount
    } catch {
        Write-Host "  [WARN] No primary stats for $Smtp : $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # Quota cap + percent used
    $QuotaGB = ConvertTo-GB $Mailbox.ProhibitSendReceiveQuota
    $QuotaDisplay = if ($null -eq $QuotaGB) { 'Unlimited' } else { $QuotaGB }
    $PercentUsed  = if ($QuotaGB -and $QuotaGB -gt 0 -and $null -ne $UsageGB) {
        [math]::Round(($UsageGB / $QuotaGB) * 100, 1)
    } else { $null }

    # Archive
    $ArchiveGuid    = "$($Mailbox.ArchiveGuid)"
    $ArchiveEnabled = ($ArchiveGuid -and $ArchiveGuid -ne $EmptyGuid)
    $ArchiveUsageGB = $null
    $ArchiveItems   = $null
    $ArchiveQuotaGB = $null
    if ($ArchiveEnabled) {
        try {
            $aStats         = Get-MailboxStatistics -Identity $Smtp -Archive -ErrorAction Stop
            $ArchiveUsageGB = ConvertTo-GB $aStats.TotalItemSize
            $ArchiveItems   = $aStats.ItemCount
        } catch {
            Write-Host "  [WARN] No archive stats for $Smtp : $($_.Exception.Message)" -ForegroundColor Yellow
        }
        $ArchiveQuotaGB = ConvertTo-GB $Mailbox.ArchiveQuota
    }

    $Results.Add([PSCustomObject]@{
        DisplayName      = $Mailbox.DisplayName
        EmailAddress     = $Smtp
        Type             = $Type
        CurrentUsageGB   = $UsageGB
        ItemCount        = $Items
        QuotaGB          = $QuotaDisplay
        PercentUsed      = $PercentUsed
        ArchiveEnabled   = if ($ArchiveEnabled) { 'Yes' } else { 'No' }
        ArchiveUsageGB   = $ArchiveUsageGB
        ArchiveItemCount = $ArchiveItems
        ArchiveQuotaGB   = $ArchiveQuotaGB
    })
}

Write-Progress -Activity "Mailbox Storage" -Completed

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

$UserCount    = ($Results | Where-Object { $_.Type -eq 'User'   }).Count
$SharedCount  = ($Results | Where-Object { $_.Type -eq 'Shared' }).Count
$ArchiveCount = ($Results | Where-Object { $_.ArchiveEnabled -eq 'Yes' }).Count
$TotalUsage   = [math]::Round((($Results | Where-Object { $null -ne $_.CurrentUsageGB } | Measure-Object -Property CurrentUsageGB -Sum).Sum), 2)
$TotalArchive = [math]::Round((($Results | Where-Object { $null -ne $_.ArchiveUsageGB } | Measure-Object -Property ArchiveUsageGB -Sum).Sum), 2)

$RunTime = "{0:D2}h {1:D2}m {2:D2}s {3:D3}ms" -f `
    $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds, $Elapsed.Milliseconds

Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |                  RUN SUMMARY                    |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Tenant Code           : {0,-25}|" -f $TenantCode)   -ForegroundColor White
Write-Host ("  | Total Mailboxes       : {0,-25}|" -f $TotalCount)   -ForegroundColor White
Write-Host ("  |   User                : {0,-25}|" -f $UserCount)    -ForegroundColor White
Write-Host ("  |   Shared              : {0,-25}|" -f $SharedCount)  -ForegroundColor White
Write-Host ("  | Archives Enabled      : {0,-25}|" -f $ArchiveCount) -ForegroundColor White
Write-Host ("  | Total Primary Usage   : {0,-25}|" -f "$TotalUsage GB")   -ForegroundColor White
Write-Host ("  | Total Archive Usage   : {0,-25}|" -f "$TotalArchive GB") -ForegroundColor White
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Total Run Time        : {0,-25}|" -f $RunTime)      -ForegroundColor Yellow
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""
Read-Host "  Press Enter to exit"
