#Requires -Modules Microsoft.Graph.Users, Microsoft.Graph.Identity.SignIns

<#
.SYNOPSIS
    Bulk-invite external users as B2B guests in a Microsoft 365 tenant.

.DESCRIPTION
    Connects to Microsoft Graph, auto-detects the tenant ID, and builds the
    invitation redirect URL as:

        https://myapplications.microsoft.com/?tenantid=<TenantID>

    The user can confirm or override this URL.

    The script then accepts a pasted list of entries, one per line:

        DisplayName, EmailAddress

    If a line contains no comma, the email is used as the display name.

    Microsoft's invitation email is always sent (SendInvitationMessage = true).

    For each existing guest with a matching email, the script prompts:
        [S]kip / [R]e-invite / Skip [A]ll remaining / [C]ancel run.

    Results are exported to C:\MSP-M365-Utility\.

.NOTES
    Required Modules:
        - Microsoft.Graph.Users
        - Microsoft.Graph.Identity.SignIns
    Required Permissions (Graph):
        - User.Invite.All
        - User.Read.All
#>

# ----------------------------------------------------------
#  CONFIGURATION
# ----------------------------------------------------------
$ErrorActionPreference = 'Stop'
$Host.UI.RawUI.WindowTitle = "M365 Bulk Guest Invite"

# ----------------------------------------------------------
#  BANNER
# ----------------------------------------------------------
Clear-Host
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |         M365 Bulk Guest Invite Script            |" -ForegroundColor Cyan
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
$OutputFile = "BulkGuestInvite_${TenantCode}_${Timestamp}.csv"
$OutputPath = Join-Path $OutputRoot $OutputFile

# ----------------------------------------------------------
#  MODULE CHECK + CONNECT
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [1/4] Checking required modules..." -ForegroundColor Cyan
foreach ($Mod in @('Microsoft.Graph.Users','Microsoft.Graph.Identity.SignIns')) {
    if (-not (Get-Module -ListAvailable -Name $Mod)) {
        Write-Host "  [!] Module '$Mod' not found. Installing..." -ForegroundColor Yellow
        Install-Module -Name $Mod -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module -Name $Mod -ErrorAction Stop
    Write-Host "  [OK] $Mod loaded." -ForegroundColor Green
}

Write-Host ""
Write-Host "  [2/4] Connecting to Microsoft Graph..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes "User.Invite.All","User.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "  [OK] Microsoft Graph connected." -ForegroundColor Green
}
catch {
    Write-Host "  [ERROR] Failed to connect: $_" -ForegroundColor Red
    Read-Host "`n  Press Enter to exit"; exit 1
}

# ----------------------------------------------------------
#  AUTO-DETECT TENANT ID + REDIRECT URL
# ----------------------------------------------------------
$TenantId = $null
try {
    $ctx = Get-MgContext
    $TenantId = $ctx.TenantId
} catch {
    Write-Host "  [ERROR] Could not read tenant context: $_" -ForegroundColor Red
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Read-Host "`n  Press Enter to exit"; exit 1
}

$DefaultUrl = "https://myapplications.microsoft.com/?tenantid=$TenantId"

Write-Host ""
Write-Host "  Tenant Code  : $TenantCode" -ForegroundColor Green
Write-Host "  Tenant ID    : $TenantId"   -ForegroundColor Green
Write-Host "  Redirect URL : $DefaultUrl" -ForegroundColor Green
Write-Host ""

$UrlConfirm = Read-Host "  Use this redirect URL? (Y/N)"
if ($UrlConfirm -in @('N','n')) {
    $custom = Read-Host "  Enter custom InviteRedirectUrl"
    if ($custom -and $custom.Trim() -ne '') {
        $InviteRedirectUrl = $custom.Trim()
    } else {
        Write-Host "  No URL provided - using default." -ForegroundColor Yellow
        $InviteRedirectUrl = $DefaultUrl
    }
} else {
    $InviteRedirectUrl = $DefaultUrl
}

Write-Host "  Using URL    : $InviteRedirectUrl" -ForegroundColor Cyan

# ----------------------------------------------------------
#  PASTE-LIST INPUT
# ----------------------------------------------------------
Write-Host ""
Write-Host "  Paste guest entries below (one per line) in the format:" -ForegroundColor Yellow
Write-Host "      DisplayName, EmailAddress" -ForegroundColor White
Write-Host "  Examples:" -ForegroundColor DarkGray
Write-Host "      Jane Doe, jane@partner.com" -ForegroundColor DarkGray
Write-Host "      vendor@external.com" -ForegroundColor DarkGray
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

    $commaIdx    = $trimmed.IndexOf(',')
    $displayName = $null
    $email       = $null

    if ($commaIdx -ge 0) {
        $displayName = $trimmed.Substring(0, $commaIdx).Trim()
        $email       = $trimmed.Substring($commaIdx + 1).Trim()
    } else {
        $email       = $trimmed
        $displayName = $trimmed
    }

    if ($email -notmatch $EmailRegex) {
        [void]$Malformed.Add($line); continue
    }

    if (-not $displayName -or $displayName -eq $email) {
        $displayName = $email
    }

    $key = $email.ToLower()
    if ($SeenEmails.ContainsKey($key)) { continue }
    $SeenEmails[$key] = $true

    [void]$Entries.Add([PSCustomObject]@{
        DisplayName  = $displayName
        EmailAddress = $email
    })
}

if ($Entries.Count -eq 0) {
    Write-Host ""
    Write-Host "  [!] No valid entries found. Nothing to do." -ForegroundColor Yellow
    Disconnect-MgGraph -ErrorAction SilentlyContinue
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
    Write-Host ("    {0,-35} {1}" -f $_.DisplayName, $_.EmailAddress)
}
if ($Malformed.Count -gt 0) {
    Write-Host ""
    Write-Host "  Skipped (malformed):" -ForegroundColor Yellow
    $Malformed | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
}

$Confirm = Read-Host "`n  Proceed with sending invitations? (Y/N)"
if ($Confirm -notin @('Y','y')) {
    Write-Host "  Aborted." -ForegroundColor Red
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Read-Host "`n  Press Enter to exit"; exit 0
}

$ScriptStart = [System.Diagnostics.Stopwatch]::StartNew()

# ----------------------------------------------------------
#  PROCESS WITH PROGRESS BAR + PER-DUPLICATE PROMPT
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [3/4] Sending invitations..." -ForegroundColor Cyan
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
        -Activity        "Bulk Guest Invite -- Tenant: $TenantCode" `
        -Status          "$ProgressBar  $Counter / $Total  ($PercentComplete%)  |  $($Entry.DisplayName)" `
        -PercentComplete  $PercentComplete `
        -CurrentOperation "Current: $($Entry.EmailAddress)"

    if ($Cancelled) {
        $Results.Add([PSCustomObject]@{
            DisplayName         = $Entry.DisplayName
            EmailAddress        = $Entry.EmailAddress
            Status              = 'Not-Run-Cancelled'
            InvitationRedeemUrl = $null
            InvitedUserId       = $null
            ErrorMessage        = $null
        })
        continue
    }

    # Check for existing guest with this email (mail OR otherMails)
    $existing = $null
    try {
        $escaped = $Entry.EmailAddress.Replace("'", "''")
        $existing = Get-MgUser -Filter "mail eq '$escaped' and userType eq 'Guest'" -ConsistencyLevel eventual -Property "Id,DisplayName,UserPrincipalName,Mail,UserType" -ErrorAction Stop | Select-Object -First 1
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
            Write-Host "  [DUPLICATE] '$($Entry.EmailAddress)' already exists as guest - '$($existing.DisplayName)' ($($existing.UserPrincipalName))." -ForegroundColor Yellow
            $answer = Read-Host "    [S]kip / [R]e-invite / Skip [A]ll remaining / [C]ancel run"
            switch ($answer.ToUpper()) {
                'R' { $action = 'ReInvite' }
                'A' { $SkipAllDupes = $true; $action = 'Skip' }
                'C' { $Cancelled = $true; $action = 'Cancel' }
                default { $action = 'Skip' }
            }
        }

        switch ($action) {
            'Skip' {
                $Results.Add([PSCustomObject]@{
                    DisplayName         = $Entry.DisplayName
                    EmailAddress        = $Entry.EmailAddress
                    Status              = 'Skipped-Exists'
                    InvitationRedeemUrl = $null
                    InvitedUserId       = $existing.Id
                    ErrorMessage        = $null
                })
                Write-Host "    Skipped." -ForegroundColor DarkGray
                continue
            }
            'Cancel' {
                $Results.Add([PSCustomObject]@{
                    DisplayName         = $Entry.DisplayName
                    EmailAddress        = $Entry.EmailAddress
                    Status              = 'Cancelled'
                    InvitationRedeemUrl = $null
                    InvitedUserId       = $existing.Id
                    ErrorMessage        = $null
                })
                Write-Host "    Run cancelled by user." -ForegroundColor Red
                continue
            }
            # 'ReInvite' falls through to the New-MgInvitation call below
        }
    }

    # Send invitation
    try {
        $invite = New-MgInvitation `
            -InvitedUserEmailAddress $Entry.EmailAddress `
            -InvitedUserDisplayName  $Entry.DisplayName `
            -InviteRedirectUrl       $InviteRedirectUrl `
            -SendInvitationMessage   `
            -ErrorAction Stop

        $statusLabel = if ($existing) { 'Re-Invited' } else { 'Invited' }
        $Results.Add([PSCustomObject]@{
            DisplayName         = $Entry.DisplayName
            EmailAddress        = $Entry.EmailAddress
            Status              = $statusLabel
            InvitationRedeemUrl = $invite.InviteRedeemUrl
            InvitedUserId       = $invite.InvitedUser.Id
            ErrorMessage        = $null
        })
        Write-Host "  $statusLabel : $($Entry.EmailAddress)" -ForegroundColor Green
    }
    catch {
        $Results.Add([PSCustomObject]@{
            DisplayName         = $Entry.DisplayName
            EmailAddress        = $Entry.EmailAddress
            Status              = 'Failed'
            InvitationRedeemUrl = $null
            InvitedUserId       = $null
            ErrorMessage        = $_.Exception.Message
        })
        Write-Host "  Failed  : $($Entry.EmailAddress) - $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Progress -Activity "Bulk Guest Invite" -Completed

# ----------------------------------------------------------
#  EXPORT RESULTS
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [4/4] Exporting results..." -ForegroundColor Cyan
try {
    $Results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "  [OK] Results CSV exported." -ForegroundColor Green
    Write-Host "       Path: $OutputPath"     -ForegroundColor DarkGray
} catch {
    Write-Host "  [ERROR] Failed to export results: $_" -ForegroundColor Red
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

$InvitedCnt    = ($Results | Where-Object { $_.Status -eq 'Invited'        }).Count
$ReInvitedCnt  = ($Results | Where-Object { $_.Status -eq 'Re-Invited'     }).Count
$SkippedCnt    = ($Results | Where-Object { $_.Status -eq 'Skipped-Exists' }).Count
$FailedCnt     = ($Results | Where-Object { $_.Status -eq 'Failed'         }).Count
$CancelledCnt  = ($Results | Where-Object { $_.Status -in @('Cancelled','Not-Run-Cancelled') }).Count

$RunTime = "{0:D2}h {1:D2}m {2:D2}s {3:D3}ms" -f `
    $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds, $Elapsed.Milliseconds

Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |                  RUN SUMMARY                    |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Tenant Code        : {0,-28}|" -f $TenantCode)   -ForegroundColor White
Write-Host ("  | Redirect URL       : {0,-28}|" -f $(if ($InviteRedirectUrl.Length -gt 28) {$InviteRedirectUrl.Substring(0,25)+'...'} else {$InviteRedirectUrl})) -ForegroundColor White
Write-Host ("  | Total Entries      : {0,-28}|" -f $Total)        -ForegroundColor White
Write-Host ("  |   Invited          : {0,-28}|" -f $InvitedCnt)   -ForegroundColor White
Write-Host ("  |   Re-Invited       : {0,-28}|" -f $ReInvitedCnt) -ForegroundColor White
Write-Host ("  |   Skipped (exists) : {0,-28}|" -f $SkippedCnt)   -ForegroundColor White
Write-Host ("  |   Failed           : {0,-28}|" -f $FailedCnt)    -ForegroundColor White
Write-Host ("  |   Cancelled        : {0,-28}|" -f $CancelledCnt) -ForegroundColor White
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Total Run Time     : {0,-28}|" -f $RunTime)      -ForegroundColor Yellow
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""
Read-Host "  Press Enter to exit"
