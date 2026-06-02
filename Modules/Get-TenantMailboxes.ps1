#Requires -Modules ExchangeOnlineManagement, Microsoft.Graph.Users

<#
.SYNOPSIS
    Retrieves all mailboxes from a Microsoft 365 tenant and exports details to CSV.

.DESCRIPTION
    Connects to Exchange Online and Microsoft Graph to gather mailbox information
    including DisplayName, EmailAddress, RecipientTypeDetails, IsEnabled, and
    IsLicensed status. Displays a real-time progress bar and exports results to CSV.

.NOTES
    Required Modules:
        - ExchangeOnlineManagement  : Install-Module ExchangeOnlineManagement
        - Microsoft.Graph.Users     : Install-Module Microsoft.Graph.Users
    Required Permissions:
        - Exchange Online : View-Only Recipients (or higher)
        - Microsoft Graph : User.Read.All
#>

# ----------------------------------------------------------
#  CONFIGURATION
# ----------------------------------------------------------
$ErrorActionPreference = 'Stop'
$Host.UI.RawUI.WindowTitle = "M365 Mailbox Inventory"

# ----------------------------------------------------------
#  BANNER
# ----------------------------------------------------------
Clear-Host
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |        M365 Mailbox Inventory Script             |" -ForegroundColor Cyan
Write-Host "  |   Exchange Online  *  Microsoft Graph            |" -ForegroundColor Cyan
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

# Derive output file name from tenant code + timestamp
$Timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputFile = "Mailboxes_${TenantCode}_${Timestamp}.csv"
$OutputPath = Join-Path $PSScriptRoot $OutputFile

Write-Host ""
Write-Host "  Tenant Code : $TenantCode"  -ForegroundColor Green
Write-Host "  Output File : $OutputPath"  -ForegroundColor Green
Write-Host ""

# ----------------------------------------------------------
#  START TIMER
# ----------------------------------------------------------
$ScriptStart = [System.Diagnostics.Stopwatch]::StartNew()

# ----------------------------------------------------------
#  MODULE CHECK
# ----------------------------------------------------------
Write-Host "  [1/4] Checking required modules..." -ForegroundColor Cyan

$RequiredModules = @('ExchangeOnlineManagement', 'Microsoft.Graph.Users')
foreach ($Mod in $RequiredModules) {
    if (-not (Get-Module -ListAvailable -Name $Mod)) {
        Write-Host "  [!] Module '$Mod' not found. Installing..." -ForegroundColor Yellow
        Install-Module -Name $Mod -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module -Name $Mod -ErrorAction Stop
    Write-Host "  [OK] $Mod loaded." -ForegroundColor Green
}

# ----------------------------------------------------------
#  CONNECT TO SERVICES
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [2/4] Connecting to Microsoft 365 services..." -ForegroundColor Cyan

try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "  [OK] Exchange Online connected." -ForegroundColor Green

    Connect-MgGraph -Scopes "User.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "  [OK] Microsoft Graph connected." -ForegroundColor Green
}
catch {
    Write-Host "  [ERROR] Failed to connect: $_" -ForegroundColor Red
    exit 1
}

# ----------------------------------------------------------
#  RETRIEVE MAILBOXES
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [3/4] Retrieving mailboxes from Exchange Online..." -ForegroundColor Cyan

try {
    $Mailboxes  = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop
    $TotalCount = $Mailboxes.Count
    Write-Host "  [OK] Found $TotalCount mailbox(es)." -ForegroundColor Green
}
catch {
    Write-Host "  [ERROR] Could not retrieve mailboxes: $_" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    exit 1
}

# Pre-fetch all Graph users into a hashtable for fast lookup (UPN -> user object)
Write-Host "  Fetching user data from Microsoft Graph..." -ForegroundColor DarkCyan

$GraphUsers = @{}
Get-MgUser -All -Property "UserPrincipalName,AccountEnabled,AssignedLicenses" |
    ForEach-Object { $GraphUsers[$_.UserPrincipalName.ToLower()] = $_ }

$GraphCount = $GraphUsers.Count
Write-Host "  [OK] Graph user data cached ($GraphCount entries)." -ForegroundColor Green

# ----------------------------------------------------------
#  PROCESS MAILBOXES WITH PROGRESS BAR
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [4/4] Processing mailboxes..." -ForegroundColor Cyan
Write-Host ""

$Results = [System.Collections.Generic.List[PSCustomObject]]::new()
$Counter = 0

foreach ($Mbx in $Mailboxes) {
    $Counter++

    # Progress Bar
    $PercentComplete = [math]::Round(($Counter / $TotalCount) * 100, 1)
    $BarWidth        = 40
    $FilledBars      = [math]::Floor($PercentComplete / 100 * $BarWidth)
    $EmptyBars       = $BarWidth - $FilledBars
    $ProgressBar     = "[" + ("#" * $FilledBars) + ("-" * $EmptyBars) + "]"

    Write-Progress `
        -Activity        "Processing Mailboxes -- Tenant: $TenantCode" `
        -Status          "$ProgressBar  $Counter / $TotalCount  ($PercentComplete%)  |  $($Mbx.DisplayName)" `
        -PercentComplete  $PercentComplete `
        -CurrentOperation "Current: $($Mbx.PrimarySmtpAddress)"

    # Graph lookup
    $UPN       = $Mbx.UserPrincipalName
    $GraphUser = $GraphUsers[$UPN.ToLower()]

    $IsEnabled  = if ($GraphUser) { $GraphUser.AccountEnabled }                   else { $null }
    $IsLicensed = if ($GraphUser) { ($GraphUser.AssignedLicenses.Count -gt 0) }   else { $null }

    # Friendly recipient type label
    $RecipientLabel = switch -Wildcard ($Mbx.RecipientTypeDetails) {
        "UserMailbox"       { "User"     }
        "SharedMailbox"     { "Shared"   }
        "RoomMailbox"       { "Resource" }
        "EquipmentMailbox"  { "Resource" }
        "SchedulingMailbox" { "Resource" }
        default             { $Mbx.RecipientTypeDetails }
    }

    $Results.Add([PSCustomObject]@{
        DisplayName          = $Mbx.DisplayName
        EmailAddress         = $Mbx.PrimarySmtpAddress
        RecipientTypeDetails = $RecipientLabel
        IsEnabled            = $IsEnabled
        IsLicensed           = $IsLicensed
    })
}

Write-Progress -Activity "Processing Mailboxes" -Completed

# ----------------------------------------------------------
#  EXPORT TO CSV
# ----------------------------------------------------------
try {
    $Results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "  [OK] CSV exported successfully." -ForegroundColor Green
    Write-Host "       Path: $OutputPath"          -ForegroundColor DarkGray
}
catch {
    Write-Host "  [ERROR] Failed to export CSV: $_" -ForegroundColor Red
}

# ----------------------------------------------------------
#  DISCONNECT
# ----------------------------------------------------------
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Disconnect-MgGraph -ErrorAction SilentlyContinue
Write-Host "  Sessions disconnected." -ForegroundColor DarkGray

# ----------------------------------------------------------
#  SUMMARY
# ----------------------------------------------------------
$ScriptStart.Stop()
$Elapsed = $ScriptStart.Elapsed

$UserCount     = ($Results | Where-Object { $_.RecipientTypeDetails -eq "User"     }).Count
$SharedCount   = ($Results | Where-Object { $_.RecipientTypeDetails -eq "Shared"   }).Count
$ResourceCount = ($Results | Where-Object { $_.RecipientTypeDetails -eq "Resource" }).Count
$OtherCount    = ($Results | Where-Object { $_.RecipientTypeDetails -notin @("User","Shared","Resource") }).Count
$EnabledCount  = ($Results | Where-Object { $_.IsEnabled  -eq $true }).Count
$LicensedCount = ($Results | Where-Object { $_.IsLicensed -eq $true }).Count

$RunTime = "{0:D2}h {1:D2}m {2:D2}s {3:D3}ms" -f `
    $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds, $Elapsed.Milliseconds

Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |                  RUN SUMMARY                    |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Tenant Code     : {0,-31}|" -f $TenantCode)    -ForegroundColor White
Write-Host ("  | Total Mailboxes : {0,-31}|" -f $TotalCount)    -ForegroundColor White
Write-Host ("  |   User          : {0,-31}|" -f $UserCount)     -ForegroundColor White
Write-Host ("  |   Shared        : {0,-31}|" -f $SharedCount)   -ForegroundColor White
Write-Host ("  |   Resource      : {0,-31}|" -f $ResourceCount) -ForegroundColor White
Write-Host ("  |   Other         : {0,-31}|" -f $OtherCount)    -ForegroundColor White
Write-Host ("  | Enabled         : {0,-31}|" -f $EnabledCount)  -ForegroundColor White
Write-Host ("  | Licensed        : {0,-31}|" -f $LicensedCount) -ForegroundColor White
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Total Run Time  : {0,-31}|" -f $RunTime)       -ForegroundColor Yellow
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""
