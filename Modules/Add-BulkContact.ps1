#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Bulk-create Exchange Online Mail Contacts (external addresses) from a
    pasted list of name + email entries.

.DESCRIPTION
    Connects to Exchange Online and creates a Mail Contact for each entry
    pasted into the console. Each line should be in the format:

        DisplayName, ExternalEmailAddress

    If a line contains no comma, the email is used as the display name.

    The script:
        - Validates each line and skips malformed entries
        - Previews the parsed list and asks for confirmation
        - Checks for existing contacts (by email) and, per duplicate,
          asks whether to Skip, Update properties, Skip All, or Cancel
        - Logs each action to a results CSV under C:\MSP-M365-Utility\

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
$Host.UI.RawUI.WindowTitle = "M365 Bulk Contact Creation"

# ----------------------------------------------------------
#  BANNER
# ----------------------------------------------------------
Clear-Host
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |        M365 Bulk Contact Creation Script         |" -ForegroundColor Cyan
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
    }
    catch {
        Write-Host "  [ERROR] Could not create output folder '$OutputRoot': $_" -ForegroundColor Red
        Read-Host "`n  Press Enter to exit"; exit 1
    }
}
$Timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputFile = "BulkContact_${TenantCode}_${Timestamp}.csv"
$OutputPath = Join-Path $OutputRoot $OutputFile

# ----------------------------------------------------------
#  PASTE-LIST INPUT
# ----------------------------------------------------------
Write-Host ""
Write-Host "  Paste contacts below (one per line) in the format:" -ForegroundColor Yellow
Write-Host "      DisplayName, ExternalEmailAddress" -ForegroundColor White
Write-Host "  Examples:" -ForegroundColor DarkGray
Write-Host "      Jane Doe, jane@external.com" -ForegroundColor DarkGray
Write-Host "      vendor@partner.com" -ForegroundColor DarkGray
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
$Entries    = [System.Collections.Generic.List[PSCustomObject]]::new()
$Malformed  = [System.Collections.Generic.List[string]]::new()
$SeenEmails = @{}

foreach ($line in $RawLines) {
    $trimmed = $line.Trim()
    if (-not $trimmed) { continue }

    # Split on FIRST comma only - so DisplayName can contain spaces/punctuation
    $commaIdx    = $trimmed.IndexOf(',')
    $displayName = $null
    $email       = $null

    if ($commaIdx -ge 0) {
        $displayName = $trimmed.Substring(0, $commaIdx).Trim()
        $email       = $trimmed.Substring($commaIdx + 1).Trim()
    } else {
        $email       = $trimmed
        $displayName = $trimmed   # placeholder; replaced below if email is valid
    }

    if ($email -notmatch $EmailRegex) {
        [void]$Malformed.Add($line)
        continue
    }

    if (-not $displayName -or $displayName -eq $email) {
        $displayName = $email
    }

    $key = $email.ToLower()
    if ($SeenEmails.ContainsKey($key)) { continue }   # in-paste dedupe
    $SeenEmails[$key] = $true

    [void]$Entries.Add([PSCustomObject]@{
        DisplayName          = $displayName
        ExternalEmailAddress = $email
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
#  PREVIEW
# ----------------------------------------------------------
Write-Host ""
Write-Host "  Parsed Entries  : $($Entries.Count)" -ForegroundColor Cyan
Write-Host "  Malformed Lines : $($Malformed.Count)" -ForegroundColor $(if ($Malformed.Count) { 'Yellow' } else { 'DarkGray' })
Write-Host ""
$Entries | ForEach-Object {
    Write-Host ("    {0,-35} {1}" -f $_.DisplayName, $_.ExternalEmailAddress)
}
if ($Malformed.Count -gt 0) {
    Write-Host ""
    Write-Host "  Skipped (malformed):" -ForegroundColor Yellow
    $Malformed | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
}

$Confirm = Read-Host "`n  Proceed with creation? (Y/N)"
if ($Confirm -notin @('Y','y')) {
    Write-Host "  Aborted." -ForegroundColor Red
    Read-Host "`n  Press Enter to exit"; exit 0
}

# ----------------------------------------------------------
#  TIMER
# ----------------------------------------------------------
$ScriptStart = [System.Diagnostics.Stopwatch]::StartNew()

# ----------------------------------------------------------
#  MODULE CHECK + CONNECT
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [1/3] Checking required modules..." -ForegroundColor Cyan
foreach ($Mod in @('ExchangeOnlineManagement')) {
    if (-not (Get-Module -ListAvailable -Name $Mod)) {
        Write-Host "  [!] Module '$Mod' not found. Installing..." -ForegroundColor Yellow
        Install-Module -Name $Mod -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module -Name $Mod -ErrorAction Stop
    Write-Host "  [OK] $Mod loaded." -ForegroundColor Green
}

Write-Host ""
Write-Host "  [2/3] Connecting to Exchange Online..." -ForegroundColor Cyan
try {
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "  [OK] Exchange Online connected." -ForegroundColor Green
}
catch {
    Write-Host "  [ERROR] Failed to connect: $_" -ForegroundColor Red
    Read-Host "`n  Press Enter to exit"; exit 1
}

# ----------------------------------------------------------
#  PROCESS WITH PROGRESS BAR + PER-DUPLICATE PROMPT
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [3/3] Processing entries..." -ForegroundColor Cyan
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
        -Activity        "Bulk Contact Creation -- Tenant: $TenantCode" `
        -Status          "$ProgressBar  $Counter / $Total  ($PercentComplete%)  |  $($Entry.DisplayName)" `
        -PercentComplete  $PercentComplete `
        -CurrentOperation "Current: $($Entry.ExternalEmailAddress)"

    if ($Cancelled) {
        $Results.Add([PSCustomObject]@{
            DisplayName          = $Entry.DisplayName
            ExternalEmailAddress = $Entry.ExternalEmailAddress
            Status               = 'Not-Run-Cancelled'
            ErrorMessage         = $null
        })
        continue
    }

    # Check for existing recipient with this email
    $existing = $null
    try {
        $existing = Get-Recipient -Identity $Entry.ExternalEmailAddress -ErrorAction Stop
    } catch {
        $existing = $null
    }

    if ($existing) {
        $action = 'Skip'
        if ($SkipAllDupes) {
            $action = 'Skip'
        }
        else {
            Write-Host ""
            Write-Host "  [DUPLICATE] '$($Entry.ExternalEmailAddress)' already exists as $($existing.RecipientTypeDetails) - '$($existing.DisplayName)'." -ForegroundColor Yellow
            $answer = Read-Host "    [S]kip / [U]pdate properties / Skip [A]ll remaining / [C]ancel run"
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
                    DisplayName          = $Entry.DisplayName
                    ExternalEmailAddress = $Entry.ExternalEmailAddress
                    Status               = 'Skipped-Duplicate'
                    ErrorMessage         = $null
                })
                Write-Host "    Skipped." -ForegroundColor DarkGray
                continue
            }
            'Update' {
                try {
                    if ($existing.RecipientTypeDetails -eq 'MailContact') {
                        Set-Contact -Identity $existing.Identity -DisplayName $Entry.DisplayName -ErrorAction Stop
                        $Results.Add([PSCustomObject]@{
                            DisplayName          = $Entry.DisplayName
                            ExternalEmailAddress = $Entry.ExternalEmailAddress
                            Status               = 'Updated'
                            ErrorMessage         = $null
                        })
                        Write-Host "    Updated DisplayName." -ForegroundColor Green
                    } else {
                        $Results.Add([PSCustomObject]@{
                            DisplayName          = $Entry.DisplayName
                            ExternalEmailAddress = $Entry.ExternalEmailAddress
                            Status               = 'Skipped-NotAContact'
                            ErrorMessage         = "Existing object is a $($existing.RecipientTypeDetails), not a MailContact - refusing to modify."
                        })
                        Write-Host "    Not a MailContact ($($existing.RecipientTypeDetails)); skipped." -ForegroundColor Yellow
                    }
                } catch {
                    $Results.Add([PSCustomObject]@{
                        DisplayName          = $Entry.DisplayName
                        ExternalEmailAddress = $Entry.ExternalEmailAddress
                        Status               = 'Failed-Update'
                        ErrorMessage         = $_.Exception.Message
                    })
                    Write-Host "    Update failed: $($_.Exception.Message)" -ForegroundColor Red
                }
                continue
            }
            'Cancel' {
                $Results.Add([PSCustomObject]@{
                    DisplayName          = $Entry.DisplayName
                    ExternalEmailAddress = $Entry.ExternalEmailAddress
                    Status               = 'Cancelled'
                    ErrorMessage         = $null
                })
                Write-Host "    Run cancelled by user." -ForegroundColor Red
                continue
            }
        }
    }

    # Brand-new contact
    try {
        New-MailContact `
            -Name                 $Entry.DisplayName `
            -DisplayName          $Entry.DisplayName `
            -ExternalEmailAddress $Entry.ExternalEmailAddress `
            -ErrorAction Stop | Out-Null

        $Results.Add([PSCustomObject]@{
            DisplayName          = $Entry.DisplayName
            ExternalEmailAddress = $Entry.ExternalEmailAddress
            Status               = 'Created'
            ErrorMessage         = $null
        })
        Write-Host "  Created : $($Entry.ExternalEmailAddress)" -ForegroundColor Green
    } catch {
        $Results.Add([PSCustomObject]@{
            DisplayName          = $Entry.DisplayName
            ExternalEmailAddress = $Entry.ExternalEmailAddress
            Status               = 'Failed'
            ErrorMessage         = $_.Exception.Message
        })
        Write-Host "  Failed  : $($Entry.ExternalEmailAddress) - $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Progress -Activity "Bulk Contact Creation" -Completed

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
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Write-Host "  Session disconnected." -ForegroundColor DarkGray

# ----------------------------------------------------------
#  SUMMARY
# ----------------------------------------------------------
$ScriptStart.Stop()
$Elapsed = $ScriptStart.Elapsed

$CreatedCnt   = ($Results | Where-Object { $_.Status -eq 'Created'           }).Count
$UpdatedCnt   = ($Results | Where-Object { $_.Status -eq 'Updated'           }).Count
$SkippedDupe  = ($Results | Where-Object { $_.Status -eq 'Skipped-Duplicate' }).Count
$SkippedOther = ($Results | Where-Object { $_.Status -eq 'Skipped-NotAContact' }).Count
$FailedCnt    = ($Results | Where-Object { $_.Status -like 'Failed*'         }).Count
$CancelledCnt = ($Results | Where-Object { $_.Status -in @('Cancelled','Not-Run-Cancelled') }).Count

$RunTime = "{0:D2}h {1:D2}m {2:D2}s {3:D3}ms" -f `
    $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds, $Elapsed.Milliseconds

Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |                  RUN SUMMARY                    |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Tenant Code         : {0,-27}|" -f $TenantCode)   -ForegroundColor White
Write-Host ("  | Total Entries       : {0,-27}|" -f $Total)        -ForegroundColor White
Write-Host ("  |   Created           : {0,-27}|" -f $CreatedCnt)   -ForegroundColor White
Write-Host ("  |   Updated           : {0,-27}|" -f $UpdatedCnt)   -ForegroundColor White
Write-Host ("  |   Skipped (dupe)    : {0,-27}|" -f $SkippedDupe)  -ForegroundColor White
Write-Host ("  |   Skipped (other)   : {0,-27}|" -f $SkippedOther) -ForegroundColor White
Write-Host ("  |   Failed            : {0,-27}|" -f $FailedCnt)    -ForegroundColor White
Write-Host ("  |   Cancelled         : {0,-27}|" -f $CancelledCnt) -ForegroundColor White
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Total Run Time      : {0,-27}|" -f $RunTime)      -ForegroundColor Yellow
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""
Read-Host "  Press Enter to exit"
