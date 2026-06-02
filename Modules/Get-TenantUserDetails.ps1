#Requires -Modules Microsoft.Graph.Users

<#
.SYNOPSIS
    Retrieves user details from a Microsoft 365 tenant and exports to CSV.

.DESCRIPTION
    Connects to Microsoft Graph to gather user information including
    FirstName, LastName, PrimarySmtpAddress, Title, Department, and Manager.
    Displays a real-time progress bar and exports results to CSV.

.NOTES
    Required Modules:
        - Microsoft.Graph.Users : Install-Module Microsoft.Graph.Users
    Required Permissions:
        - Microsoft Graph : User.Read.All
                           OrgContact.Read.All (for Manager lookup)
#>

# ----------------------------------------------------------
#  CONFIGURATION
# ----------------------------------------------------------
$ErrorActionPreference = 'Stop'
$Host.UI.RawUI.WindowTitle = "M365 User Details Inventory"

# ----------------------------------------------------------
#  BANNER
# ----------------------------------------------------------
Clear-Host
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |        M365 User Details Inventory Script        |" -ForegroundColor Cyan
Write-Host "  |              Microsoft Graph                     |" -ForegroundColor Cyan
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
$OutputFile = "UserDetails_${TenantCode}_${Timestamp}.csv"
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

$RequiredModules = @('Microsoft.Graph.Users')
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
Write-Host "  [2/4] Connecting to Microsoft Graph..." -ForegroundColor Cyan

try {
    Connect-MgGraph -Scopes "User.Read.All","OrgContact.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "  [OK] Microsoft Graph connected." -ForegroundColor Green
}
catch {
    Write-Host "  [ERROR] Failed to connect: $_" -ForegroundColor Red
    Read-Host "`n  Press Enter to exit"
    exit 1
}

# ----------------------------------------------------------
#  RETRIEVE USERS
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [3/4] Retrieving users from Microsoft Graph..." -ForegroundColor Cyan

try {
    $Users = Get-MgUser -All `
        -Property "GivenName,Surname,Mail,JobTitle,Department,UserPrincipalName,Id" `
        -ErrorAction Stop

    $TotalCount = $Users.Count
    Write-Host "  [OK] Found $TotalCount user(s)." -ForegroundColor Green
}
catch {
    Write-Host "  [ERROR] Could not retrieve users: $_" -ForegroundColor Red
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Read-Host "`n  Press Enter to exit"
    exit 1
}

# ----------------------------------------------------------
#  PROCESS USERS WITH PROGRESS BAR
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [4/4] Processing users..." -ForegroundColor Cyan
Write-Host ""

$Results = [System.Collections.Generic.List[PSCustomObject]]::new()
$Counter = 0

foreach ($User in $Users) {
    $Counter++

    # Progress Bar
    $PercentComplete = [math]::Round(($Counter / $TotalCount) * 100, 1)
    $BarWidth        = 40
    $FilledBars      = [math]::Floor($PercentComplete / 100 * $BarWidth)
    $EmptyBars       = $BarWidth - $FilledBars
    $ProgressBar     = "[" + ("#" * $FilledBars) + ("-" * $EmptyBars) + "]"

    $DisplayName = "$($User.GivenName) $($User.Surname)".Trim()

    Write-Progress `
        -Activity        "Processing Users -- Tenant: $TenantCode" `
        -Status          "$ProgressBar  $Counter / $TotalCount  ($PercentComplete%)  |  $DisplayName" `
        -PercentComplete  $PercentComplete `
        -CurrentOperation "Current: $(if ($User.Mail) { $User.Mail } else { $User.UserPrincipalName })"

    # Manager lookup (not all users have a manager - handle gracefully)
    $ManagerName = $null
    try {
        $ManagerObj  = Get-MgUserManager -UserId $User.Id -ErrorAction Stop
        $ManagerName = $ManagerObj.AdditionalProperties['displayName']
    }
    catch {
        $ManagerName = $null
    }

    $Results.Add([PSCustomObject]@{
        FirstName          = $User.GivenName
        LastName           = $User.Surname
        PrimarySmtpAddress = if ($User.Mail) { $User.Mail } else { $User.UserPrincipalName }
        Title              = $User.JobTitle
        Department         = $User.Department
        Manager            = $ManagerName
    })
}

Write-Progress -Activity "Processing Users" -Completed

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
Disconnect-MgGraph -ErrorAction SilentlyContinue
Write-Host "  Session disconnected." -ForegroundColor DarkGray

# ----------------------------------------------------------
#  SUMMARY
# ----------------------------------------------------------
$ScriptStart.Stop()
$Elapsed = $ScriptStart.Elapsed

$WithTitleCount      = ($Results | Where-Object { $_.Title      }).Count
$WithDeptCount       = ($Results | Where-Object { $_.Department }).Count
$WithManagerCount    = ($Results | Where-Object { $_.Manager    }).Count
$WithSmtpCount       = ($Results | Where-Object { $_.PrimarySmtpAddress }).Count

$RunTime = "{0:D2}h {1:D2}m {2:D2}s {3:D3}ms" -f `
    $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds, $Elapsed.Milliseconds

Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |                  RUN SUMMARY                    |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Tenant Code      : {0,-30}|" -f $TenantCode)         -ForegroundColor White
Write-Host ("  | Total Users      : {0,-30}|" -f $TotalCount)         -ForegroundColor White
Write-Host ("  | With SMTP        : {0,-30}|" -f $WithSmtpCount)      -ForegroundColor White
Write-Host ("  | With Title       : {0,-30}|" -f $WithTitleCount)     -ForegroundColor White
Write-Host ("  | With Department  : {0,-30}|" -f $WithDeptCount)      -ForegroundColor White
Write-Host ("  | With Manager     : {0,-30}|" -f $WithManagerCount)   -ForegroundColor White
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Total Run Time   : {0,-30}|" -f $RunTime)            -ForegroundColor Yellow
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""
Read-Host "  Press Enter to exit"
