#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Grant calendar permissions on a target mailbox to one or more users,
    with a per-user choice of Reviewer / Author / Editor.

.DESCRIPTION
    Connects to Exchange Online and adds a folder permission on the target
    mailbox's Calendar folder for each user in the pasted list.

    Paste format (one entry per line):

        user@domain.com                  - role will be prompted later
        user@domain.com, R               - Reviewer
        user@domain.com, Author          - Author
        user@domain.com, E               - Editor

    Supported role aliases:
        Reviewer  : R  | Reviewer  (read-only)
        Author    : A  | Author    (edit own items)
        Editor    : E  | Editor    (read / write all)

    Behaviour:
        - The Calendar folder name is resolved per-mailbox so this works in
          non-English tenants too.
        - When a user already has a permission, the script prompts
          [S]kip / [U]pdate to the new role / Skip [A]ll / [C]ancel.
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
$Host.UI.RawUI.WindowTitle = "M365 Add Calendar Permission"

# ----------------------------------------------------------
#  BANNER
# ----------------------------------------------------------
Clear-Host
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |        M365 Add Calendar Permission Script       |" -ForegroundColor Cyan
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
#  INPUT - TARGET MAILBOX (whose calendar)
# ----------------------------------------------------------
$EmailRegex = '^[^@\s]+@[^@\s]+\.[^@\s]+$'
do {
    $TargetMailbox = (Read-Host "  Enter the target MAILBOX (whose calendar to share)").Trim()
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
$OutputFile = "CalendarPermAdd_${TenantCode}_${Timestamp}.csv"
$OutputPath = Join-Path $OutputRoot $OutputFile

# ----------------------------------------------------------
#  PASTE-LIST INPUT
# ----------------------------------------------------------
Write-Host ""
Write-Host "  Paste user entries below (one per line):" -ForegroundColor Yellow
Write-Host "      user@domain.com              (role will be prompted)" -ForegroundColor DarkGray
Write-Host "      user@domain.com, R           (Reviewer)" -ForegroundColor DarkGray
Write-Host "      user@domain.com, Author      (Author)" -ForegroundColor DarkGray
Write-Host "      user@domain.com, E           (Editor)" -ForegroundColor DarkGray
Write-Host "  When finished, press ENTER on a blank line." -ForegroundColor Yellow
Write-Host ""

$RawLines = @()
while ($true) {
    $line = Read-Host
    if ($line -eq "") { break }
    $RawLines += $line
}

# ----------------------------------------------------------
#  HELPER - normalise role aliases to canonical names
# ----------------------------------------------------------
function ConvertTo-Role {
    param ([string]$Raw)
    if (-not $Raw) { return $null }
    switch ($Raw.Trim().ToUpper()) {
        'R'        { 'Reviewer'; break }
        'REVIEWER' { 'Reviewer'; break }
        'A'        { 'Author';   break }
        'AUTHOR'   { 'Author';   break }
        'E'        { 'Editor';   break }
        'EDITOR'   { 'Editor';   break }
        default    { $null }
    }
}

# ----------------------------------------------------------
#  PARSE & VALIDATE
# ----------------------------------------------------------
$Entries    = [System.Collections.Generic.List[PSCustomObject]]::new()
$Malformed  = [System.Collections.Generic.List[string]]::new()
$SeenEmails = @{}

foreach ($line in $RawLines) {
    $trimmed = $line.Trim()
    if (-not $trimmed) { continue }

    $parts = $trimmed.Split(',', 2)
    $email = $parts[0].Trim()
    $role  = if ($parts.Count -gt 1) { ConvertTo-Role $parts[1] } else { $null }

    if ($email -notmatch $EmailRegex) {
        [void]$Malformed.Add($line); continue
    }
    if ($parts.Count -gt 1 -and -not $role) {
        # role given but unrecognised - mark malformed so user knows
        [void]$Malformed.Add("$line  (unknown role: '$($parts[1].Trim())')"); continue
    }

    $key = $email.ToLower()
    if ($SeenEmails.ContainsKey($key)) { continue }
    $SeenEmails[$key] = $true

    [void]$Entries.Add([PSCustomObject]@{
        UserEmail = $email
        Role      = $role
    })
}

if ($Entries.Count -eq 0) {
    Write-Host ""
    Write-Host "  [!] No valid entries found. Nothing to do." -ForegroundColor Yellow
    if ($Malformed.Count -gt 0) {
        Write-Host "  Malformed lines:" -ForegroundColor DarkGray
        $Malformed | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
    }
    Read-Host "`n  Press Enter to exit"; exit 0
}

# ----------------------------------------------------------
#  FILL IN MISSING ROLES (per-user prompt)
# ----------------------------------------------------------
$WithoutRole = @($Entries | Where-Object { -not $_.Role })
if ($WithoutRole.Count -gt 0) {
    Write-Host ""
    Write-Host "  $($WithoutRole.Count) entry(ies) have no role - choose one for each:" -ForegroundColor Cyan
    foreach ($entry in $WithoutRole) {
        do {
            $resp = Read-Host "    $($entry.UserEmail)  ->  [R]eviewer / [A]uthor / [E]ditor"
            $entry.Role = ConvertTo-Role $resp
            if (-not $entry.Role) {
                Write-Host "      [!] Enter R, A, or E (or Reviewer / Author / Editor)." -ForegroundColor Yellow
            }
        } while (-not $entry.Role)
    }
}

# ----------------------------------------------------------
#  PREVIEW
# ----------------------------------------------------------
Write-Host ""
Write-Host "  Target Mailbox  : $TargetMailbox" -ForegroundColor Cyan
Write-Host "  Parsed Entries  : $($Entries.Count)" -ForegroundColor Cyan
Write-Host "  Malformed Lines : $($Malformed.Count)" -ForegroundColor $(if ($Malformed.Count) { 'Yellow' } else { 'DarkGray' })
Write-Host ""
$Entries | ForEach-Object {
    Write-Host ("    {0,-40} ->  {1}" -f $_.UserEmail, $_.Role)
}
if ($Malformed.Count -gt 0) {
    Write-Host ""
    Write-Host "  Skipped (malformed):" -ForegroundColor Yellow
    $Malformed | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
}

$Confirm = Read-Host "`n  Proceed with granting permissions? (Y/N)"
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
#  RESOLVE CALENDAR FOLDER PATH (localization-aware)
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
#  PROCESS WITH PROGRESS BAR + PER-DUPLICATE PROMPT
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [4/4] Granting permissions..." -ForegroundColor Cyan
Write-Host ""

$Results       = [System.Collections.Generic.List[PSCustomObject]]::new()
$SkipAllDupes  = $false
$Cancelled     = $false
$Total         = $Entries.Count
$Counter       = 0

foreach ($Entry in $Entries) {
    $Counter++
    $PercentComplete = [math]::Round(($Counter / $Total) * 100, 1)
    $BarWidth        = 40
    $FilledBars      = [math]::Floor($PercentComplete / 100 * $BarWidth)
    $EmptyBars       = $BarWidth - $FilledBars
    $ProgressBar     = "[" + ("#" * $FilledBars) + ("-" * $EmptyBars) + "]"

    Write-Progress `
        -Activity        "Add Calendar Permission -- Tenant: $TenantCode" `
        -Status          "$ProgressBar  $Counter / $Total  ($PercentComplete%)  |  $($Entry.UserEmail)" `
        -PercentComplete  $PercentComplete `
        -CurrentOperation "Role: $($Entry.Role)"

    if ($Cancelled) {
        $Results.Add([PSCustomObject]@{
            UserEmail = $Entry.UserEmail; Role = $Entry.Role
            Status = 'Not-Run-Cancelled'; ErrorMessage = $null
        })
        continue
    }

    # Check existing permission for this user on the target's calendar
    $existing = $null
    try {
        $existing = Get-MailboxFolderPermission -Identity $CalendarPath -User $Entry.UserEmail -ErrorAction Stop
    } catch {
        $existing = $null   # Treat any error here as "no existing permission"
    }

    if ($existing) {
        $action = 'Skip'
        if ($SkipAllDupes) { $action = 'Skip' }
        else {
            Write-Host ""
            Write-Host "  [DUPLICATE] '$($Entry.UserEmail)' already has '$($existing.AccessRights -join ',')' on $TargetMailbox's calendar." -ForegroundColor Yellow
            $answer = Read-Host "    [S]kip / [U]pdate to '$($Entry.Role)' / Skip [A]ll remaining / [C]ancel run"
            switch ($answer.ToUpper()) {
                'U' { $action = 'Update' }
                'A' { $SkipAllDupes = $true; $action = 'Skip' }
                'C' { $Cancelled = $true; $action = 'Cancel' }
                default { $action = 'Skip' }
            }
        }

        switch ($action) {
            'Skip' {
                $Results.Add([PSCustomObject]@{
                    UserEmail = $Entry.UserEmail; Role = $Entry.Role
                    Status = 'Skipped-AlreadyHasAccess'
                    ErrorMessage = "Existing access: $($existing.AccessRights -join ',')"
                })
                Write-Host "    Skipped." -ForegroundColor DarkGray
                continue
            }
            'Update' {
                try {
                    Set-MailboxFolderPermission -Identity $CalendarPath -User $Entry.UserEmail -AccessRights $Entry.Role -ErrorAction Stop
                    $Results.Add([PSCustomObject]@{
                        UserEmail = $Entry.UserEmail; Role = $Entry.Role
                        Status = 'Updated'; ErrorMessage = $null
                    })
                    Write-Host "    Updated to $($Entry.Role)." -ForegroundColor Green
                } catch {
                    $Results.Add([PSCustomObject]@{
                        UserEmail = $Entry.UserEmail; Role = $Entry.Role
                        Status = 'Failed-Update'; ErrorMessage = $_.Exception.Message
                    })
                    Write-Host "    Update failed: $($_.Exception.Message)" -ForegroundColor Red
                }
                continue
            }
            'Cancel' {
                $Results.Add([PSCustomObject]@{
                    UserEmail = $Entry.UserEmail; Role = $Entry.Role
                    Status = 'Cancelled'; ErrorMessage = $null
                })
                Write-Host "    Run cancelled by user." -ForegroundColor Red
                continue
            }
        }
    }

    # No existing permission - add fresh
    try {
        Add-MailboxFolderPermission -Identity $CalendarPath -User $Entry.UserEmail -AccessRights $Entry.Role -ErrorAction Stop | Out-Null
        $Results.Add([PSCustomObject]@{
            UserEmail = $Entry.UserEmail; Role = $Entry.Role
            Status = 'Granted'; ErrorMessage = $null
        })
        Write-Host "  Granted : $($Entry.UserEmail)  ($($Entry.Role))" -ForegroundColor Green
    } catch {
        $Results.Add([PSCustomObject]@{
            UserEmail = $Entry.UserEmail; Role = $Entry.Role
            Status = 'Failed'; ErrorMessage = $_.Exception.Message
        })
        Write-Host "  Failed  : $($Entry.UserEmail) - $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Progress -Activity "Add Calendar Permission" -Completed

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

$GrantedCnt   = ($Results | Where-Object { $_.Status -eq 'Granted'           }).Count
$UpdatedCnt   = ($Results | Where-Object { $_.Status -eq 'Updated'           }).Count
$SkippedCnt   = ($Results | Where-Object { $_.Status -eq 'Skipped-AlreadyHasAccess' }).Count
$FailedCnt    = ($Results | Where-Object { $_.Status -like 'Failed*'         }).Count
$CancelledCnt = ($Results | Where-Object { $_.Status -in @('Cancelled','Not-Run-Cancelled') }).Count

$RunTime = "{0:D2}h {1:D2}m {2:D2}s {3:D3}ms" -f `
    $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds, $Elapsed.Milliseconds

Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |                  RUN SUMMARY                    |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Tenant Code         : {0,-27}|" -f $TenantCode)   -ForegroundColor White
Write-Host ("  | Target Mailbox      : {0,-27}|" -f $TargetMailbox)-ForegroundColor White
Write-Host ("  | Total Entries       : {0,-27}|" -f $Total)        -ForegroundColor White
Write-Host ("  |   Granted           : {0,-27}|" -f $GrantedCnt)   -ForegroundColor White
Write-Host ("  |   Updated           : {0,-27}|" -f $UpdatedCnt)   -ForegroundColor White
Write-Host ("  |   Skipped (dupe)    : {0,-27}|" -f $SkippedCnt)   -ForegroundColor White
Write-Host ("  |   Failed            : {0,-27}|" -f $FailedCnt)    -ForegroundColor White
Write-Host ("  |   Cancelled         : {0,-27}|" -f $CancelledCnt) -ForegroundColor White
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Total Run Time      : {0,-27}|" -f $RunTime)      -ForegroundColor Yellow
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""
Read-Host "  Press Enter to exit"
