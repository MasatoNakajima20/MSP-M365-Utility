#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Bulk-grant FullAccess and SendAs on a target mailbox (or create the
    target as a shared mailbox first if it does not exist).

.DESCRIPTION
    Connects to Exchange Online, ensures the target mailbox exists (creating
    it as a Shared Mailbox if not), then grants FullAccess and SendAs to
    every user in the pasted list.

    For each user:
        - If they already have both rights  -> Skipped-AlreadyHasBoth
        - Add only what they're missing      -> Granted-Both / Granted-FullAccessOnly / Granted-SendAsOnly
        - Failures are recorded per right    -> Failed-FullAccess / Failed-SendAs / Failed-Both

    Results are written to C:\MSP-M365-Utility\.

.NOTES
    Required Modules:
        - ExchangeOnlineManagement
    Required Permissions (EXO):
        - Recipient Management (for permission cmdlets)
        - Mail Recipient Creation (only if a shared mailbox needs creating)
#>

# ----------------------------------------------------------
#  CONFIGURATION
# ----------------------------------------------------------
$ErrorActionPreference = 'Stop'
$Host.UI.RawUI.WindowTitle = "M365 Bulk Add Mailbox Access"

# ----------------------------------------------------------
#  BANNER
# ----------------------------------------------------------
Clear-Host
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |     M365 Bulk Add Mailbox Access Script          |" -ForegroundColor Cyan
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

# Derive default display name from local part (used only if creation is needed)
$LocalPart          = ($TargetMailbox -split '@')[0]
$DefaultDisplayName = (
    ($LocalPart -replace '[._\-]', ' ' -split ' ' |
        Where-Object { $_ } |
        ForEach-Object { (Get-Culture).TextInfo.ToTitleCase($_.ToLower()) }
    ) -join ' '
)

$DisplayName = (Read-Host "  Display Name (only used if creating new shared mailbox) [$DefaultDisplayName]").Trim()
if (-not $DisplayName) { $DisplayName = $DefaultDisplayName }

# ----------------------------------------------------------
#  INPUT - AUTO-MAPPING
# ----------------------------------------------------------
$AutoMappingAnswer = (Read-Host "  Enable Outlook Auto-Mapping for FullAccess? (Y/N) [Y]").Trim()
$AutoMapping      = -not ($AutoMappingAnswer -in @('N','n'))

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
$OutputFile = "MailboxAccessAdd_${TenantCode}_${Timestamp}.csv"
$OutputPath = Join-Path $OutputRoot $OutputFile

# ----------------------------------------------------------
#  PASTE-LIST INPUT
# ----------------------------------------------------------
Write-Host ""
Write-Host "  Paste user emails below (one per line)." -ForegroundColor Yellow
Write-Host "  Each user will receive BOTH FullAccess and SendAs on $TargetMailbox." -ForegroundColor DarkGray
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
Write-Host "  [1/5] Checking required modules..." -ForegroundColor Cyan
foreach ($Mod in @('ExchangeOnlineManagement')) {
    if (-not (Get-Module -ListAvailable -Name $Mod)) {
        Write-Host "  [!] Module '$Mod' not found. Installing..." -ForegroundColor Yellow
        Install-Module -Name $Mod -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module -Name $Mod -ErrorAction Stop
    Write-Host "  [OK] $Mod loaded." -ForegroundColor Green
}

Write-Host ""
Write-Host "  [2/5] Connecting to Exchange Online..." -ForegroundColor Cyan
try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "  [OK] Exchange Online connected." -ForegroundColor Green
} catch {
    Write-Host "  [ERROR] Failed to connect: $_" -ForegroundColor Red
    Read-Host "`n  Press Enter to exit"; exit 1
}

# ----------------------------------------------------------
#  CHECK IF TARGET EXISTS
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [3/5] Checking target mailbox..." -ForegroundColor Cyan
$existingMbx = $null
try {
    $existingMbx = Get-Mailbox -Identity $TargetMailbox -ErrorAction Stop
} catch {
    $existingMbx = $null
}

$WillCreate = -not $existingMbx
if ($existingMbx) {
    Write-Host "  [OK] Found: $($existingMbx.DisplayName) ($($existingMbx.RecipientTypeDetails))" -ForegroundColor Green
} else {
    Write-Host "  [!]  Target not found - will be created as Shared Mailbox." -ForegroundColor Yellow
}

# ----------------------------------------------------------
#  PREVIEW
# ----------------------------------------------------------
Write-Host ""
Write-Host "  Tenant Code      : $TenantCode" -ForegroundColor Cyan
Write-Host "  Target Mailbox   : $TargetMailbox" -ForegroundColor Cyan
if ($WillCreate) {
    Write-Host "  Action           : CREATE as Shared Mailbox" -ForegroundColor Yellow
    Write-Host "  Display Name     : $DisplayName" -ForegroundColor Cyan
    Write-Host "  Name (alias)     : $LocalPart" -ForegroundColor Cyan
} else {
    Write-Host "  Action           : USE existing ($($existingMbx.RecipientTypeDetails))" -ForegroundColor Cyan
    Write-Host "  Existing Display : $($existingMbx.DisplayName)" -ForegroundColor Cyan
}
Write-Host "  Auto-Mapping     : $(if ($AutoMapping) {'Enabled'} else {'Disabled'})" -ForegroundColor Cyan
Write-Host "  Users to Grant   : $($Users.Count)" -ForegroundColor Cyan
Write-Host "  Malformed Lines  : $($Malformed.Count)" -ForegroundColor $(if ($Malformed.Count) { 'Yellow' } else { 'DarkGray' })
Write-Host ""
$Users | ForEach-Object { Write-Host "    $_" }
if ($Malformed.Count -gt 0) {
    Write-Host ""
    Write-Host "  Skipped (malformed):" -ForegroundColor Yellow
    $Malformed | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
}

$Confirm = Read-Host "`n  Proceed? (Y/N)"
if ($Confirm -notin @('Y','y')) {
    Write-Host "  Aborted." -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Read-Host "`n  Press Enter to exit"; exit 0
}

$ScriptStart = [System.Diagnostics.Stopwatch]::StartNew()

# ----------------------------------------------------------
#  CREATE SHARED MAILBOX (if needed)
# ----------------------------------------------------------
if ($WillCreate) {
    Write-Host ""
    Write-Host "  [4/5] Creating shared mailbox..." -ForegroundColor Cyan
    try {
        New-Mailbox -Shared `
            -Name                $LocalPart `
            -DisplayName         $DisplayName `
            -PrimarySmtpAddress  $TargetMailbox `
            -ErrorAction Stop | Out-Null
        Write-Host "  [OK] Shared mailbox created." -ForegroundColor Green

        # Provisioning can take a moment - poll until Get-Mailbox returns it
        $waitedSec = 0
        while ($waitedSec -lt 30) {
            try {
                $existingMbx = Get-Mailbox -Identity $TargetMailbox -ErrorAction Stop
                if ($existingMbx) { break }
            } catch { }
            Start-Sleep -Seconds 3
            $waitedSec += 3
        }
        if (-not $existingMbx) {
            Write-Host "  [WARN] New mailbox not yet visible after $waitedSec s. Continuing anyway." -ForegroundColor Yellow
        } else {
            Write-Host "  [OK] Mailbox is queryable." -ForegroundColor Green
        }
    } catch {
        Write-Host "  [ERROR] Failed to create shared mailbox: $_" -ForegroundColor Red
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Read-Host "`n  Press Enter to exit"; exit 1
    }
} else {
    Write-Host ""
    Write-Host "  [4/5] Skipping create - mailbox already exists." -ForegroundColor DarkGray
}

# ----------------------------------------------------------
#  GRANT PERMISSIONS PER USER
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [5/5] Granting permissions..." -ForegroundColor Cyan
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
        -Activity        "Add Mailbox Access -- Tenant: $TenantCode" `
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

    if ($hasFA -and $hasSA) {
        $Results.Add([PSCustomObject]@{
            UserEmail = $User; GrantedFullAccess = $false; GrantedSendAs = $false
            Status = 'Skipped-AlreadyHasBoth'; ErrorMessage = $null
        })
        Write-Host "  Skipped : $User  (already has FullAccess + SendAs)" -ForegroundColor DarkGray
        continue
    }

    $faOk      = $hasFA
    $saOk      = $hasSA
    $faError   = $null
    $saError   = $null

    if (-not $hasFA) {
        try {
            Add-MailboxPermission -Identity $TargetMailbox -User $User -AccessRights FullAccess -AutoMapping:$AutoMapping -ErrorAction Stop | Out-Null
            $faOk = $true
        } catch {
            $faError = $_.Exception.Message
        }
    }

    if (-not $hasSA) {
        try {
            Add-RecipientPermission -Identity $TargetMailbox -Trustee $User -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
            $saOk = $true
        } catch {
            $saError = $_.Exception.Message
        }
    }

    # Derive status
    $status = $null
    $msg    = $null
    if ($faOk -and $saOk) {
        if ($hasFA -and -not $hasSA)       { $status = 'Granted-SendAsOnly' }
        elseif (-not $hasFA -and $hasSA)   { $status = 'Granted-FullAccessOnly' }
        else                               { $status = 'Granted-Both' }
    }
    elseif (-not $faOk -and -not $saOk)    { $status = 'Failed-Both';      $msg = "FA: $faError | SA: $saError" }
    elseif ($faOk -and -not $saOk)         { $status = 'Failed-SendAs';    $msg = $saError }
    elseif (-not $faOk -and $saOk)         { $status = 'Failed-FullAccess';$msg = $faError }

    $Results.Add([PSCustomObject]@{
        UserEmail         = $User
        GrantedFullAccess = ($faOk -and -not $hasFA)
        GrantedSendAs     = ($saOk -and -not $hasSA)
        Status            = $status
        ErrorMessage      = $msg
    })

    $colour = if ($status -like 'Granted-*') { 'Green' } elseif ($status -like 'Failed-*') { 'Red' } else { 'Yellow' }
    Write-Host "  $status : $User" -ForegroundColor $colour
}

Write-Progress -Activity "Add Mailbox Access" -Completed

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

$BothCnt        = ($Results | Where-Object { $_.Status -eq 'Granted-Both'           }).Count
$FAOnlyCnt      = ($Results | Where-Object { $_.Status -eq 'Granted-FullAccessOnly' }).Count
$SAOnlyCnt      = ($Results | Where-Object { $_.Status -eq 'Granted-SendAsOnly'     }).Count
$SkippedCnt     = ($Results | Where-Object { $_.Status -eq 'Skipped-AlreadyHasBoth' }).Count
$FailedAnyCnt   = ($Results | Where-Object { $_.Status -like 'Failed-*'             }).Count

$RunTime = "{0:D2}h {1:D2}m {2:D2}s {3:D3}ms" -f `
    $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds, $Elapsed.Milliseconds

Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |                  RUN SUMMARY                    |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Tenant Code         : {0,-27}|" -f $TenantCode)    -ForegroundColor White
Write-Host ("  | Target Mailbox      : {0,-27}|" -f $TargetMailbox) -ForegroundColor White
Write-Host ("  | Created Mailbox     : {0,-27}|" -f $(if ($WillCreate) {'Yes'} else {'No'})) -ForegroundColor White
Write-Host ("  | Auto-Mapping        : {0,-27}|" -f $(if ($AutoMapping) {'Enabled'} else {'Disabled'})) -ForegroundColor White
Write-Host ("  | Total Users         : {0,-27}|" -f $Total)         -ForegroundColor White
Write-Host ("  |   Granted Both      : {0,-27}|" -f $BothCnt)       -ForegroundColor White
Write-Host ("  |   Granted FA only   : {0,-27}|" -f $FAOnlyCnt)     -ForegroundColor White
Write-Host ("  |   Granted SA only   : {0,-27}|" -f $SAOnlyCnt)     -ForegroundColor White
Write-Host ("  |   Skipped (had both): {0,-27}|" -f $SkippedCnt)    -ForegroundColor White
Write-Host ("  |   Failed (any)      : {0,-27}|" -f $FailedAnyCnt)  -ForegroundColor White
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Total Run Time      : {0,-27}|" -f $RunTime)       -ForegroundColor Yellow
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""
Read-Host "  Press Enter to exit"
