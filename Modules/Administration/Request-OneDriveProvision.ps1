#Requires -Modules Microsoft.Online.SharePoint.PowerShell

<#
.SYNOPSIS
    Pre-provision OneDrive for Business personal sites for a pasted list of
    users, processed one at a time.

.DESCRIPTION
    Connects to SharePoint Online and calls Request-SPOPersonalSite for each
    user email in the pasted list, one by one, so every request is logged
    individually (you can see exactly which users succeeded or failed).

    Paste format (one per line):

        user@domain.com

    Notes:
        - Requesting per-user (rather than in batches) is slower but gives a
          clean per-user result. A short pause between requests keeps SPO from
          throttling.
        - Pre-provisioning only queues the OneDrive site (-NoWait); the site
          itself may take a few minutes to finish creating on Microsoft's side.
        - Users must be licensed for OneDrive / SharePoint or the request fails.

    Results are written to C:\MSP-M365-Utility\.

.NOTES
    Required Modules:
        - Microsoft.Online.SharePoint.PowerShell
          (Install-Module Microsoft.Online.SharePoint.PowerShell)
    Required Permissions:
        - SharePoint Administrator (or Global Administrator)
    Compatibility:
        - The SharePoint Online Management Shell is most reliable under
          Windows PowerShell 5.1. If a cmdlet misbehaves under PowerShell 7,
          run this module in Windows PowerShell instead.
#>

# ----------------------------------------------------------
#  CONFIGURATION
# ----------------------------------------------------------
$ErrorActionPreference = 'Stop'
$Host.UI.RawUI.WindowTitle = "M365 OneDrive Pre-Provisioning"

# ----------------------------------------------------------
#  BANNER
# ----------------------------------------------------------
Clear-Host
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |       M365 OneDrive Pre-Provisioning Script      |" -ForegroundColor Cyan
Write-Host "  |              SharePoint Online                   |" -ForegroundColor Cyan
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
#  INPUT - SHAREPOINT ADMIN URL
# ----------------------------------------------------------
do {
    Write-Host ""
    Write-Host "  Enter the SharePoint ADMIN URL (e.g. https://contoso-admin.sharepoint.com)" -ForegroundColor Yellow
    $SharePointAdminUrl = (Read-Host "  SharePoint Admin URL").Trim()

    $looksValid = $SharePointAdminUrl -match '^https://[^/\s]+\.sharepoint\.com/?$'
    if (-not $looksValid) {
        Write-Host "  [!] That doesn't look like a SharePoint admin URL." -ForegroundColor Yellow
    }
    elseif ($SharePointAdminUrl -notmatch '-admin\.sharepoint\.com') {
        Write-Host "  [!] Tip: the admin URL usually contains '-admin' (e.g. contoso-admin.sharepoint.com)." -ForegroundColor DarkYellow
        $useAnyway = Read-Host "      Use this URL anyway? (Y/N)"
        if ($useAnyway -notin @('Y','y')) { $looksValid = $false }
    }
} while (-not $looksValid)

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
$OutputFile = "OneDriveProvision_${TenantCode}_${Timestamp}.csv"
$OutputPath = Join-Path $OutputRoot $OutputFile

# ----------------------------------------------------------
#  PASTE-LIST INPUT
# ----------------------------------------------------------
Write-Host ""
Write-Host "  Paste user emails below (one per line)." -ForegroundColor Yellow
Write-Host "  Each user's OneDrive will be pre-provisioned, one at a time." -ForegroundColor DarkGray
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
$EmailRegex = '^[^@\s]+@[^@\s]+\.[^@\s]+$'
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
#  PREVIEW
# ----------------------------------------------------------
Write-Host ""
Write-Host "  Tenant Code        : $TenantCode" -ForegroundColor Cyan
Write-Host "  SharePoint Admin   : $SharePointAdminUrl" -ForegroundColor Cyan
Write-Host "  Users to Provision : $($Users.Count)" -ForegroundColor Cyan
Write-Host "  Malformed Lines    : $($Malformed.Count)" -ForegroundColor $(if ($Malformed.Count) { 'Yellow' } else { 'DarkGray' })
Write-Host ""
$Users | ForEach-Object { Write-Host "    $_" }
if ($Malformed.Count -gt 0) {
    Write-Host ""
    Write-Host "  Skipped (malformed):" -ForegroundColor Yellow
    $Malformed | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
}

$Confirm = Read-Host "`n  Proceed with OneDrive pre-provisioning? (Y/N)"
if ($Confirm -notin @('Y','y')) {
    Write-Host "  Aborted." -ForegroundColor Red
    Read-Host "`n  Press Enter to exit"; exit 0
}

$ScriptStart = [System.Diagnostics.Stopwatch]::StartNew()

# ----------------------------------------------------------
#  MODULE CHECK + CONNECT
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [1/3] Checking required modules..." -ForegroundColor Cyan
foreach ($Mod in @('Microsoft.Online.SharePoint.PowerShell')) {
    if (-not (Get-Module -ListAvailable -Name $Mod)) {
        Write-Host "  [!] Module '$Mod' not found. Installing..." -ForegroundColor Yellow
        Install-Module -Name $Mod -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module -Name $Mod -ErrorAction Stop
    Write-Host "  [OK] $Mod loaded." -ForegroundColor Green
}

Write-Host ""
Write-Host "  [2/3] Connecting to SharePoint Online..." -ForegroundColor Cyan
try {
    Connect-SPOService -Url $SharePointAdminUrl -ErrorAction Stop
    Write-Host "  [OK] SharePoint Online connected." -ForegroundColor Green
} catch {
    Write-Host "  [ERROR] Failed to connect: $_" -ForegroundColor Red
    Read-Host "`n  Press Enter to exit"; exit 1
}

# ----------------------------------------------------------
#  PROCESS - ONE BY ONE WITH PROGRESS BAR
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [3/3] Requesting OneDrive pre-provisioning..." -ForegroundColor Cyan
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
        -Activity        "OneDrive Pre-Provisioning -- Tenant: $TenantCode" `
        -Status          "$ProgressBar  $Counter / $Total  ($PercentComplete%)  |  $User" `
        -PercentComplete  $PercentComplete `
        -CurrentOperation "Requesting personal site..."

    try {
        Request-SPOPersonalSite -UserEmails @($User) -NoWait -ErrorAction Stop | Out-Null
        $Results.Add([PSCustomObject]@{
            UserEmail    = $User
            Status       = 'Requested'
            ErrorMessage = $null
        })
        Write-Host "  Requested : $User" -ForegroundColor Green
    } catch {
        $Results.Add([PSCustomObject]@{
            UserEmail    = $User
            Status       = 'Failed'
            ErrorMessage = $_.Exception.Message
        })
        Write-Host "  Failed    : $User - $($_.Exception.Message)" -ForegroundColor Red
    }

    # Gentle throttle between per-user requests
    Start-Sleep -Milliseconds 655
}

Write-Progress -Activity "OneDrive Pre-Provisioning" -Completed

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

# ----------------------------------------------------------
#  DISCONNECT
# ----------------------------------------------------------
Disconnect-SPOService -ErrorAction SilentlyContinue
Write-Host "  Session disconnected." -ForegroundColor DarkGray

# ----------------------------------------------------------
#  SUMMARY
# ----------------------------------------------------------
$ScriptStart.Stop()
$Elapsed = $ScriptStart.Elapsed

$RequestedCnt = ($Results | Where-Object { $_.Status -eq 'Requested' }).Count
$FailedCnt    = ($Results | Where-Object { $_.Status -eq 'Failed'    }).Count

$RunTime = "{0:D2}h {1:D2}m {2:D2}s {3:D3}ms" -f `
    $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds, $Elapsed.Milliseconds

Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |                  RUN SUMMARY                    |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Tenant Code         : {0,-27}|" -f $TenantCode)   -ForegroundColor White
Write-Host ("  | Total Users         : {0,-27}|" -f $Total)        -ForegroundColor White
Write-Host ("  |   Requested         : {0,-27}|" -f $RequestedCnt) -ForegroundColor White
Write-Host ("  |   Failed            : {0,-27}|" -f $FailedCnt)    -ForegroundColor White
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Total Run Time      : {0,-27}|" -f $RunTime)      -ForegroundColor Yellow
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Note: provisioning was queued (-NoWait). Sites may take a few" -ForegroundColor DarkGray
Write-Host "  minutes to finish creating on Microsoft's side." -ForegroundColor DarkGray
Write-Host ""
Read-Host "  Press Enter to exit"
