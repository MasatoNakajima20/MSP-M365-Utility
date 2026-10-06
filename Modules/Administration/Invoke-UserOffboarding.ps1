#Requires -Modules Microsoft.Graph.Users, Microsoft.Graph.Users.Actions, Microsoft.Graph.Identity.SignIns, Microsoft.Graph.Groups, ExchangeOnlineManagement

<#
.SYNOPSIS
    Bulk user offboarding pipeline for Microsoft 365 / Entra.

.DESCRIPTION
    For each pasted user email, runs this ordered pipeline and logs the
    result of every step per user:

        1. Disable the account
        2. Remove manager
        3. Revoke sign-in sessions
        4. Remove all MFA / authentication methods (password excluded)
        5. Revoke tokens (same Graph revoke call as step 3)
        6. Convert the mailbox to a Shared Mailbox (verified afterwards)
        6a. (Optional, prompted per account) Set forwarding to a given address
        6b. (Optional, prompted per account) Grant FullAccess + SendAs to
            one or more delegates on the mailbox
        6c. Hide the mailbox from the Global Address List (always on)
        7. Remove group memberships (DL, Mail-Enabled Security, M365, Security)
           and the user's access to shared mailboxes (FullAccess + SendAs)
        8. Check mailbox size (50 GB cap)
        9. Remove licenses - ONLY when it is safe to do so:
              - no mailbox exists                      -> remove (safe)
              - mailbox converted and <= 50 GB         -> remove
              - mailbox conversion failed / unverified -> SKIP + flag (retained)
              - mailbox size lookup failed             -> SKIP + flag (retained)
              - mailbox > 50 GB                         -> SKIP + flag (retained)
           Licenses are removed one SKU at a time, so a group-assigned SKU that
           cannot be removed is logged as failed without blocking the direct ones.

    Connects Microsoft Graph FIRST, then Exchange Online, to avoid the MSAL
    "WithLogging" assembly conflict.

    Two confirmation gates: (1) type the exact account count to start the run,
    then (2) a per-account Y/N prompt immediately before each account is
    touched. Answering anything but Y skips that account (logged as Skipped).

    Every action is logged. Successful actions are shown with a blank Detail;
    the Detail is only populated for actions that need attention (Partial /
    Info / Skipped / Failed). A timestamped run log is written to
    C:\Logging\MSP-M365-Utility\ and the results CSV is written to
    C:\MSP-M365-Utility\. A summary is printed at the end (including the
    license-retained flags).

.NOTES
    Required Modules:
        - Microsoft.Graph.Users
        - Microsoft.Graph.Users.Actions
        - Microsoft.Graph.Identity.SignIns
        - Microsoft.Graph.Groups
        - ExchangeOnlineManagement
    Required Permissions:
        - Microsoft Graph : User.ReadWrite.All, UserAuthenticationMethod.ReadWrite.All,
                            Group.ReadWrite.All, GroupMember.ReadWrite.All, Directory.Read.All
        - Exchange Online : Recipient Management (convert to shared, permissions)

    Output:
        - Run log : C:\Logging\MSP-M365-Utility\COMPUTER_yyyyMMdd_HHmmss_UserOffboarding.log
        - Results : C:\MSP-M365-Utility\COMPUTER_yyyyMMdd_HHmmss_UserOffboarding_<Tenant>.csv
#>

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------
$ErrorActionPreference = 'Stop'
$Host.UI.RawUI.WindowTitle = "M365 User Offboarding"
$SizeCapGB = 50   # unlicensed shared mailbox limit
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# ---------------------------------------------------------------------------
# LOGGING SETUP
# ---------------------------------------------------------------------------
# Run log goes to C:\Logging\<Project>\ per the standard; the results CSV stays
# in C:\MSP-M365-Utility\ so the launcher's "View Results" button still finds it.
$LogRoot = 'C:\Logging\MSP-M365-Utility'
if (-not (Test-Path $LogRoot)) {
    try {
        New-Item -ItemType Directory -Path $LogRoot -Force -ErrorAction Stop | Out-Null
    } catch {
        Write-Host "  [ERROR] Could not create log folder '$LogRoot': $_" -ForegroundColor Red
    }
}
$LogFile = "${env:COMPUTERNAME}_${Timestamp}_UserOffboarding.log"
$LogPath = Join-Path $LogRoot $LogFile

# Write-Log - writes a timestamped INFO/WARN/ERROR entry to the log file and
# (unless -NoConsole) to the console in a level-appropriate colour.
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO',
        [switch]$NoConsole
    )
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "$stamp [$Level] $Message"
    try { Add-Content -Path $LogPath -Value $entry -Encoding UTF8 -ErrorAction Stop } catch { }
    if (-not $NoConsole) {
        $col = switch ($Level) { 'WARN' { 'Yellow' } 'ERROR' { 'Red' } default { 'Gray' } }
        Write-Host "  $entry" -ForegroundColor $col
    }
}

# ---------------------------------------------------------------------------
# BANNER
# ---------------------------------------------------------------------------
Clear-Host
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |            M365 User Offboarding Script          |" -ForegroundColor Cyan
Write-Host "  |        Microsoft Graph  *  Exchange Online       |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""

# ---------------------------------------------------------------------------
# INPUT - TENANT CODE
# ---------------------------------------------------------------------------
do {
    $TenantCode = (Read-Host "  Enter the three-letter Tenant Code (e.g. ABC)").Trim().ToUpper()
    if ($TenantCode -notmatch '^[A-Z]{3}$') {
        Write-Host "  [!] Invalid input. Please enter exactly 3 letters (A-Z)." -ForegroundColor Yellow
    }
} while ($TenantCode -notmatch '^[A-Z]{3}$')

# ---------------------------------------------------------------------------
# OUTPUT PATH (results CSV)
# ---------------------------------------------------------------------------
$OutputRoot = 'C:\MSP-M365-Utility'
if (-not (Test-Path $OutputRoot)) {
    try {
        New-Item -ItemType Directory -Path $OutputRoot -Force -ErrorAction Stop | Out-Null
    } catch {
        Write-Host "  [ERROR] Could not create output folder '$OutputRoot': $_" -ForegroundColor Red
        Read-Host "`n  Press Enter to exit"; exit 1
    }
}
$OutputFile = "${env:COMPUTERNAME}_${Timestamp}_UserOffboarding_${TenantCode}.csv"
$OutputPath = Join-Path $OutputRoot $OutputFile

Write-Log "User Offboarding run started. Tenant: $TenantCode. Log: $LogPath" 'INFO' -NoConsole

# ---------------------------------------------------------------------------
# PASTE-LIST INPUT
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "  Paste user emails to OFFBOARD (one per line)." -ForegroundColor Yellow
Write-Host "  When finished, press ENTER on a blank line." -ForegroundColor Yellow
Write-Host ""

$RawLines = @()
while ($true) {
    $line = Read-Host
    if ($line -eq "") { break }
    $RawLines += $line
}

# ---------------------------------------------------------------------------
# PARSE AND VALIDATE
# ---------------------------------------------------------------------------
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
        if ($email -notmatch $EmailRegex) { [void]$Malformed.Add($email); continue }
        $key = $email.ToLower()
        if ($SeenEmails.ContainsKey($key)) { continue }
        $SeenEmails[$key] = $true
        [void]$Users.Add($email)
    }
}

if ($Users.Count -eq 0) {
    Write-Host ""
    Write-Host "  [!] No valid emails found. Nothing to do." -ForegroundColor Yellow
    if ($Malformed.Count -gt 0) { $Malformed | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray } }
    Write-Log "No valid emails supplied. Exiting." 'WARN' -NoConsole
    Read-Host "`n  Press Enter to exit"; exit 0
}

# ---------------------------------------------------------------------------
# PREVIEW AND STRONG CONFIRMATION (type the count)
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Red
Write-Host "  |  WARNING - THIS PERMANENTLY OFFBOARDS ACCOUNTS   |" -ForegroundColor Red
Write-Host "  +--------------------------------------------------+" -ForegroundColor Red
Write-Host "  Each account below will be:" -ForegroundColor Yellow
Write-Host "    - Disabled, manager removed, sessions/tokens revoked" -ForegroundColor DarkGray
Write-Host "    - All MFA methods removed" -ForegroundColor DarkGray
Write-Host "    - Converted to a Shared Mailbox" -ForegroundColor DarkGray
Write-Host "    - Removed from all groups and shared-mailbox access" -ForegroundColor DarkGray
Write-Host "    - License removed (unless mailbox > $SizeCapGB GB or unsafe)" -ForegroundColor DarkGray
Write-Host ""
Write-Host "  Tenant Code : $TenantCode" -ForegroundColor Cyan
Write-Host "  Accounts    : $($Users.Count)" -ForegroundColor Cyan
Write-Host ""
$Users | ForEach-Object { Write-Host "    $_" }
if ($Malformed.Count -gt 0) {
    Write-Host ""
    Write-Host "  Skipped (malformed):" -ForegroundColor Yellow
    $Malformed | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
}

Write-Host ""
$typed = Read-Host "  To PROCEED, type the number of accounts ($($Users.Count)). Anything else cancels"
if ($typed.Trim() -ne "$($Users.Count)") {
    Write-Host "  Cancelled." -ForegroundColor Red
    Write-Log "Run cancelled at the count-confirmation gate." 'WARN' -NoConsole
    Read-Host "`n  Press Enter to exit"; exit 0
}

$ScriptStart = [System.Diagnostics.Stopwatch]::StartNew()

# ---------------------------------------------------------------------------
# MODULE CHECK AND CONNECT (Graph first, then EXO)
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "  [1/3] Checking required modules..." -ForegroundColor Cyan
$RequiredModules = @(
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Users.Actions',
    'Microsoft.Graph.Identity.SignIns',
    'Microsoft.Graph.Groups',
    'ExchangeOnlineManagement'
)
foreach ($Mod in $RequiredModules) {
    if (-not (Get-Module -ListAvailable -Name $Mod)) {
        Write-Log "Module '$Mod' not found. Installing..." 'WARN'
        Install-Module -Name $Mod -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module -Name $Mod -ErrorAction Stop
    Write-Log "Module loaded: $Mod" 'INFO'
}

Write-Host ""
Write-Host "  [2/3] Connecting (Graph first, then Exchange Online)..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes "User.ReadWrite.All","UserAuthenticationMethod.ReadWrite.All","Group.ReadWrite.All","GroupMember.ReadWrite.All","Directory.Read.All" -NoWelcome -ErrorAction Stop
    Write-Log "Microsoft Graph connected." 'INFO'
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Log "Exchange Online connected." 'INFO'
} catch {
    Write-Log "Failed to connect: $($_.Exception.Message)" 'ERROR'
    Read-Host "`n  Press Enter to exit"; exit 1
}

# Build a SkuId -> friendly name map once (no extra module needed)
$SkuMap = @{}
try {
    $skuResp = Invoke-MgGraphRequest -Method GET -Uri 'v1.0/subscribedSkus' -ErrorAction Stop
    foreach ($s in $skuResp.value) { $SkuMap[$s.skuId] = $s.skuPartNumber }
} catch {
    Write-Log "Could not build SKU name map (license names will show GUIDs): $($_.Exception.Message)" 'WARN'
}

# ---------------------------------------------------------------------------
# FULLACCESS INDEX (built once, not per user)
# ---------------------------------------------------------------------------
# Enumerate every shared mailbox and its non-inherited FullAccess grantees a
# single time, keyed by grantee (lower-case). Each user's shared-mailbox
# FullAccess removals are then a hashtable lookup instead of an O(users x
# mailboxes) scan. $FullAccessIndex stays $null if the build fails, so step 7b
# can report the FullAccess check as unavailable instead of a false Success.
Write-Host ""
Write-Host "  [3/3] Building shared-mailbox FullAccess index..." -ForegroundColor Cyan
$FullAccessIndex = @{}
try {
    $allShared = @(Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop)
    foreach ($sm in $allShared) {
        try {
            $perms = @(Get-MailboxPermission -Identity $sm.Identity -ErrorAction Stop |
                       Where-Object { $_.AccessRights -contains 'FullAccess' -and -not $_.IsInherited -and "$($_.User)" -notlike 'NT AUTHORITY\*' })
            # Store the primary SMTP (fallback to Identity) - valid for removal and
            # readable when logged as the mailbox the user is removed from.
            $mbxLabel = if ($sm.PrimarySmtpAddress) { "$($sm.PrimarySmtpAddress)" } else { "$($sm.Identity)" }
            foreach ($perm in $perms) {
                $userKey = "$($perm.User)".ToLower()
                if (-not $FullAccessIndex.ContainsKey($userKey)) {
                    $FullAccessIndex[$userKey] = [System.Collections.Generic.List[string]]::new()
                }
                [void]$FullAccessIndex[$userKey].Add($mbxLabel)
            }
        } catch {
            Write-Log "FullAccess index: could not read permissions on $($sm.PrimarySmtpAddress): $($_.Exception.Message)" 'WARN' -NoConsole
        }
    }
    Write-Log "FullAccess index built over $($allShared.Count) shared mailbox(es)." 'INFO'
} catch {
    $FullAccessIndex = $null
    Write-Log "FullAccess index build failed - shared-mailbox FullAccess removal will be flagged per user: $($_.Exception.Message)" 'WARN'
}

# ---------------------------------------------------------------------------
# RESULT LOGGING HELPERS
# ---------------------------------------------------------------------------
$Results = [System.Collections.Generic.List[PSCustomObject]]::new()

# Add-Result - records one per-user action row, colours it on the console, and
# mirrors it to the run log. Success rows carry a blank Detail by design, unless
# -KeepDetail is set (used where the Detail itself is the record, e.g. the list
# of groups / shared mailboxes a user was removed from).
function Add-Result {
    param([string]$Upn, [string]$Display, [string]$Action, [string]$Status, [string]$Detail, [switch]$KeepDetail)

    # Success is shown but carries no note - the Detail is reserved for actions
    # that need attention (Partial / Info / Skipped / Failed). -KeepDetail opts a
    # row out of that blanking so its name list survives on a Success.
    if ($Status -eq 'Success' -and -not $KeepDetail) { $Detail = '' }

    $Results.Add([PSCustomObject]@{
        UserPrincipalName = $Upn
        DisplayName       = $Display
        Action            = $Action
        Status            = $Status
        Detail            = $Detail
    })
    $col = switch ($Status) {
        'Success' { 'Green' }
        'Skipped' { 'DarkGray' }
        'Info'    { 'Cyan' }
        'Partial' { 'Yellow' }
        default   { 'Red' }
    }
    Write-Host ("      {0,-26} {1,-9} {2}" -f $Action, $Status, $Detail) -ForegroundColor $col

    # Mirror to the run log (file only - the console line above is enough).
    $logLevel = switch ($Status) { 'Failed' { 'ERROR' } 'Partial' { 'WARN' } default { 'INFO' } }
    Write-Log ("{0} | {1} | {2} | {3}" -f $Upn, $Action, $Status, $Detail) $logLevel -NoConsole
}

# ConvertTo-GB - normalises an Exchange size value (string or typed) to GB as a
# number, returning $null for Unlimited or unparseable values.
function ConvertTo-GB {
    param($Size)
    if ($null -eq $Size) { return $null }
    $s = $Size.ToString()
    if ($s -match 'Unlimited') { return $null }
    if ($s -match '\(([\d,]+)\s*bytes\)') { return [math]::Round((([double]($Matches[1] -replace ',', '')) / 1GB), 2) }
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

# ---------------------------------------------------------------------------
# PROCESS EACH USER
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "  Offboarding users..." -ForegroundColor Cyan

$Total          = $Users.Count
$Counter        = 0
$ProcessedCount = 0
$SkippedCount   = 0
$NotFoundCount  = 0

foreach ($Upn in $Users) {
    $Counter++
    $PercentComplete = [math]::Round(($Counter / $Total) * 100, 1)
    Write-Progress -Activity "Offboarding -- Tenant: $TenantCode" `
                   -Status "$Counter / $Total  ($PercentComplete%)  |  $Upn" `
                   -PercentComplete $PercentComplete

    Write-Host ""
    Write-Host "  === $Upn  ($Counter of $Total) ===" -ForegroundColor White

    # Per-account confirmation - stricter than the batch gate; skip on anything but Y
    $goUser = Read-Host "      Offboard THIS account? (Y/N)"
    if ($goUser -notin @('Y','y')) {
        $SkippedCount++
        Add-Result $Upn '' 'Offboarding' 'Skipped' 'Skipped by operator at per-account prompt'
        continue
    }

    # --- Optional per-account extras (collected now, applied after conversion) ---
    # Forwarding
    $FwdTarget = $null
    $fwdAns = Read-Host "      Enable forwarding on this mailbox? (Y/N)"
    if ($fwdAns -in @('Y','y')) {
        do {
            $FwdTarget = (Read-Host "        Forward to (email address)").Trim()
            if ($FwdTarget -notmatch $EmailRegex) { Write-Host "        [!] Not a valid email address." -ForegroundColor Yellow }
        } while ($FwdTarget -notmatch $EmailRegex)
    }

    # Delegates (FullAccess + SendAs)
    $Delegates = [System.Collections.Generic.List[string]]::new()
    $permAns = Read-Host "      Add mailbox delegates (FullAccess + SendAs)? (Y/N)"
    if ($permAns -in @('Y','y')) {
        Write-Host "        Paste delegate emails (one per line). Blank line to finish:" -ForegroundColor Yellow
        $seenDel = @{}
        while ($true) {
            $dl = Read-Host
            if ($dl -eq '') { break }
            foreach ($cand in ($dl -split '[,;\s]+')) {
                $de = $cand.Trim()
                if (-not $de) { continue }
                if ($de -notmatch $EmailRegex) { Write-Host "        [!] Skipped invalid: $de" -ForegroundColor DarkGray; continue }
                $k = $de.ToLower()
                if ($seenDel.ContainsKey($k)) { continue }
                $seenDel[$k] = $true
                [void]$Delegates.Add($de)
            }
        }
    }

    # Resolve user
    $u = $null
    try {
        $u = Get-MgUser -UserId $Upn -Property "Id,DisplayName,UserPrincipalName,AccountEnabled,AssignedLicenses" -ErrorAction Stop
    } catch {
        $NotFoundCount++
        Add-Result $Upn '' 'Resolve User' 'Failed' "User not found in Graph: $($_.Exception.Message)"
        continue
    }
    $Id      = $u.Id
    $Display = $u.DisplayName
    $ProcessedCount++

    # 1) Disable account
    try {
        Update-MgUser -UserId $Id -AccountEnabled:$false -ErrorAction Stop
        Add-Result $Upn $Display 'Disable Account' 'Success' 'AccountEnabled = false'
    } catch { Add-Result $Upn $Display 'Disable Account' 'Failed' $_.Exception.Message }

    # 2) Remove manager
    try {
        Remove-MgUserManagerByRef -UserId $Id -ErrorAction Stop
        Add-Result $Upn $Display 'Remove Manager' 'Success' 'Manager reference removed'
    } catch {
        if ("$($_.Exception.Message)" -match 'ResourceNotFound|does not exist|Request_ResourceNotFound') {
            Add-Result $Upn $Display 'Remove Manager' 'Skipped' 'No manager set'
        } else {
            Add-Result $Upn $Display 'Remove Manager' 'Failed' $_.Exception.Message
        }
    }

    # 3) Revoke sign-in sessions
    try {
        Revoke-MgUserSignInSession -UserId $Id -ErrorAction Stop | Out-Null
        Add-Result $Upn $Display 'Revoke Sessions' 'Success' 'Sessions revoked'
    } catch { Add-Result $Upn $Display 'Revoke Sessions' 'Failed' $_.Exception.Message }

    # 4) Remove MFA / authentication methods
    $mfaRemoved = 0; $mfaSkipped = 0; $mfaFailed = @()
    try {
        $methods = @(Get-MgUserAuthenticationMethod -UserId $Id -ErrorAction Stop)
        foreach ($m in $methods) {
            $t = "$($m.AdditionalProperties['@odata.type'])"
            try {
                switch ($t) {
                    '#microsoft.graph.phoneAuthenticationMethod'                   { Remove-MgUserAuthenticationPhoneMethod -UserId $Id -PhoneAuthenticationMethodId $m.Id -ErrorAction Stop; $mfaRemoved++ }
                    '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod'  { Remove-MgUserAuthenticationMicrosoftAuthenticatorMethod -UserId $Id -MicrosoftAuthenticatorAuthenticationMethodId $m.Id -ErrorAction Stop; $mfaRemoved++ }
                    '#microsoft.graph.softwareOathAuthenticationMethod'            { Remove-MgUserAuthenticationSoftwareOathMethod -UserId $Id -SoftwareOathAuthenticationMethodId $m.Id -ErrorAction Stop; $mfaRemoved++ }
                    '#microsoft.graph.fido2AuthenticationMethod'                   { Remove-MgUserAuthenticationFido2Method -UserId $Id -Fido2AuthenticationMethodId $m.Id -ErrorAction Stop; $mfaRemoved++ }
                    '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod' { Remove-MgUserAuthenticationWindowsHelloForBusinessMethod -UserId $Id -WindowsHelloForBusinessAuthenticationMethodId $m.Id -ErrorAction Stop; $mfaRemoved++ }
                    '#microsoft.graph.emailAuthenticationMethod'                   { Remove-MgUserAuthenticationEmailMethod -UserId $Id -EmailAuthenticationMethodId $m.Id -ErrorAction Stop; $mfaRemoved++ }
                    '#microsoft.graph.temporaryAccessPassAuthenticationMethod'     { Remove-MgUserAuthenticationTemporaryAccessPassMethod -UserId $Id -TemporaryAccessPassAuthenticationMethodId $m.Id -ErrorAction Stop; $mfaRemoved++ }
                    '#microsoft.graph.passwordAuthenticationMethod'                { $mfaSkipped++ }   # password can't be removed
                    default                                                        { $mfaSkipped++ }
                }
            } catch { $mfaFailed += ($t -replace '#microsoft.graph.', '') }
        }
        if ($mfaFailed.Count -eq 0) {
            Add-Result $Upn $Display 'Remove MFA Methods' 'Success' "Removed $mfaRemoved; $mfaSkipped not removable (e.g. password)"
        } elseif ($mfaRemoved -gt 0) {
            Add-Result $Upn $Display 'Remove MFA Methods' 'Partial' "Removed $mfaRemoved; Failed: $($mfaFailed -join ', ')"
        } else {
            Add-Result $Upn $Display 'Remove MFA Methods' 'Failed' "Failed: $($mfaFailed -join ', ')"
        }
    } catch {
        Add-Result $Upn $Display 'Remove MFA Methods' 'Failed' $_.Exception.Message
    }

    # 5) Revoke tokens (same Graph revoke call - invalidates refresh tokens)
    try {
        Revoke-MgUserSignInSession -UserId $Id -ErrorAction Stop | Out-Null
        Add-Result $Upn $Display 'Revoke Tokens' 'Success' 'Refresh tokens invalidated'
    } catch { Add-Result $Upn $Display 'Revoke Tokens' 'Failed' $_.Exception.Message }

    # 6) Convert to Shared Mailbox
    # $MailboxType is the mailbox's RecipientTypeDetails ($null = no mailbox at
    # all). $ConvertedToShared is $true only once the mailbox is confirmed to be
    # a SharedMailbox. These two drive the license-safety decision in step 9.
    $MailboxType       = $null
    $ConvertedToShared = $false

    # Find out whether the user has a mailbox and its current type.
    try {
        $mbx = Get-Mailbox -Identity $Upn -ErrorAction Stop
        $MailboxType = "$($mbx.RecipientTypeDetails)"
    } catch {
        # Distinguish "no mailbox" (safe to remove licenses) from a real failure.
        if ("$($_.Exception.Message)" -match "couldn't be found|not found|ManagementObjectNotFound") {
            $MailboxType = $null
            Add-Result $Upn $Display 'Convert to Shared' 'Skipped' 'No mailbox found for this user'
        } else {
            $MailboxType = 'Unknown'
            Write-Log "Convert to Shared - mailbox lookup failed for ${Upn}: $($_.Exception.Message)" 'ERROR'
            Add-Result $Upn $Display 'Convert to Shared' 'Failed' "Mailbox lookup failed: $($_.Exception.Message) - license removal will be SKIPPED (RETAINED for safety)"
        }
    }

    if ($MailboxType -eq 'SharedMailbox') {
        # Already shared - treat as success, no conversion needed.
        $ConvertedToShared = $true
        Add-Result $Upn $Display 'Convert to Shared' 'Success' 'Mailbox already Shared'
    } elseif ($MailboxType -and $MailboxType -ne 'Unknown') {
        # Step 1: run the conversion and check whether the command itself errors.
        $convertErrored = $false
        try {
            Set-Mailbox -Identity $Upn -Type Shared -ErrorAction Stop
        } catch {
            $convertErrored = $true
            Write-Log "Convert to Shared failed for ${Upn}: $($_.Exception.Message)" 'ERROR'
            Add-Result $Upn $Display 'Convert to Shared' 'Failed' "$($_.Exception.Message) - license removal will be SKIPPED (RETAINED for safety)"
        }

        # Step 2: no error - poll until EXO reports the new type. The conversion
        # command returns before the type change has replicated, so an immediate
        # read can still show the old type; re-check every $PollSeconds up to
        # $MaxWaitSeconds and pass as soon as it reads SharedMailbox.
        if (-not $convertErrored) {
            $PollSeconds    = 10
            $MaxWaitSeconds = 60
            $Waited         = 0
            while ($Waited -lt $MaxWaitSeconds -and -not $ConvertedToShared) {
                Start-Sleep -Seconds $PollSeconds
                $Waited += $PollSeconds
                try {
                    $verify = Get-Mailbox -Identity $Upn -ErrorAction Stop
                    $MailboxType = "$($verify.RecipientTypeDetails)"
                    if ($MailboxType -eq 'SharedMailbox') { $ConvertedToShared = $true }
                } catch {
                    Write-Log "Convert to Shared - verify read failed for $Upn (will keep polling): $($_.Exception.Message)" 'WARN' -NoConsole
                }
            }
            if ($ConvertedToShared) {
                Add-Result $Upn $Display 'Convert to Shared' 'Success' "Mailbox type = Shared (confirmed after ${Waited}s)"
            } else {
                Add-Result $Upn $Display 'Convert to Shared' 'Failed' "Type still '$MailboxType' after ${MaxWaitSeconds}s - license removal will be SKIPPED (RETAINED for safety)"
            }
        }
    }

    # 6a) Set forwarding (only if requested at the per-account prompt)
    if ($FwdTarget) {
        try {
            Set-Mailbox -Identity $Upn -ForwardingSMTPAddress $FwdTarget -DeliverToMailboxAndForward $true -ErrorAction Stop
            Add-Result $Upn $Display 'Set Forwarding' 'Success' ''
        } catch {
            Add-Result $Upn $Display 'Set Forwarding' 'Failed' "-> $FwdTarget : $($_.Exception.Message)"
        }
    }

    # 6b) Add mailbox delegates - FullAccess + SendAs (only if requested)
    if ($Delegates.Count -gt 0) {
        $delOk = @(); $delFail = @()
        foreach ($del in $Delegates) {
            $faOk = $false; $saOk = $false; $errParts = @()
            try {
                Add-MailboxPermission -Identity $Upn -User $del -AccessRights FullAccess -AutoMapping:$true -ErrorAction Stop | Out-Null
                $faOk = $true
            } catch { $errParts += "FA: $($_.Exception.Message)" }
            try {
                Add-RecipientPermission -Identity $Upn -Trustee $del -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                $saOk = $true
            } catch { $errParts += "SA: $($_.Exception.Message)" }

            if ($faOk -and $saOk) { $delOk += $del } else { $delFail += "$del ($($errParts -join ' | '))" }
        }
        if ($delFail.Count -eq 0) {
            Add-Result $Upn $Display 'Add Mailbox Delegates' 'Success' ''
        } elseif ($delOk.Count -gt 0) {
            Add-Result $Upn $Display 'Add Mailbox Delegates' 'Partial' ("Granted: " + ($delOk -join ', ') + " | Failed: " + ($delFail -join '; '))
        } else {
            Add-Result $Upn $Display 'Add Mailbox Delegates' 'Failed' ("Failed: " + ($delFail -join '; '))
        }
    }

    # 6c) Hide the mailbox from the Global Address List (always on for offboarding)
    if ($MailboxType -and $MailboxType -ne 'Unknown') {
        try {
            Set-Mailbox -Identity $Upn -HiddenFromAddressListsEnabled $true -ErrorAction Stop
            Add-Result $Upn $Display 'Hide from GAL' 'Success' ''
        } catch {
            Add-Result $Upn $Display 'Hide from GAL' 'Failed' $_.Exception.Message
        }
    } else {
        Add-Result $Upn $Display 'Hide from GAL' 'Skipped' 'No accessible mailbox'
    }

    # 7a) Remove group memberships
    $grpRemoved = 0; $grpFailed = @(); $grpDynamic = @(); $grpRemovedNames = @()
    try {
        $memberships = @(Get-MgUserMemberOf -UserId $Id -All -ErrorAction Stop)
        foreach ($mo in $memberships) {
            $p = $mo.AdditionalProperties
            if ("$($p['@odata.type'])" -ne '#microsoft.graph.group') { continue }   # skip directory roles etc.

            $gName  = $p['displayName']
            $gMail  = $p['mail']
            $gTypes = @($p['groupTypes'])
            $mailEn = [bool]$p['mailEnabled']
            $secEn  = [bool]$p['securityEnabled']

            if ($gTypes -contains 'DynamicMembership') { $grpDynamic += $gName; continue }

            $isUnified = $gTypes -contains 'Unified'
            $useExo    = ($mailEn -and -not $isUnified)   # DL or Mail-Enabled Security -> EXO

            # Tag each group by type so the logged name shows what it was.
            $typeLabel = if ($isUnified) { 'M365' }
                         elseif ($mailEn -and $secEn) { 'Mail-Enabled Security' }
                         elseif ($mailEn) { 'DL' }
                         elseif ($secEn) { 'Security' }
                         else { 'Group' }

            try {
                if ($useExo) {
                    $ident = if ($gMail) { $gMail } else { $gName }
                    # -BypassSecurityGroupManagerCheck: mail-enabled security groups
                    # otherwise reject removal unless the operator owns the group.
                    Remove-DistributionGroupMember -Identity $ident -Member $Upn -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop
                } else {
                    Remove-MgGroupMemberByRef -GroupId $mo.Id -DirectoryObjectId $Id -ErrorAction Stop
                }
                $grpRemoved++
                $grpRemovedNames += "$gName [$typeLabel]"
            } catch { $grpFailed += "$gName ($($_.Exception.Message))" }
        }
        $detailParts = @()
        if ($grpRemovedNames.Count -gt 0) { $detailParts += ("Removed: " + ($grpRemovedNames -join ', ')) } else { $detailParts += "Removed: 0" }
        if ($grpDynamic.Count -gt 0) { $detailParts += "Dynamic (skipped): $($grpDynamic -join ', ')" }
        if ($grpFailed.Count  -gt 0) { $detailParts += "Failed: $($grpFailed -join '; ')" }
        $grpStatus = if ($grpFailed.Count -gt 0) { if ($grpRemoved -gt 0) { 'Partial' } else { 'Failed' } } else { 'Success' }
        Add-Result $Upn $Display 'Remove Group Memberships' $grpStatus ($detailParts -join ' | ') -KeepDetail
    } catch {
        Add-Result $Upn $Display 'Remove Group Memberships' 'Failed' $_.Exception.Message
    }

    # 7b) Remove shared-mailbox access (SendAs org-wide; FullAccess via the index)
    $saRemovedNames = @(); $faRemovedNames = @(); $accFailed = @()
    # SendAs - a lookup failure here is recorded, not swallowed.
    try {
        $sendAs = @(Get-RecipientPermission -Trustee $Upn -ResultSize Unlimited -ErrorAction Stop |
                    Where-Object { $_.AccessRights -contains 'SendAs' })
        foreach ($r in $sendAs) {
            try {
                Remove-RecipientPermission -Identity $r.Identity -Trustee $Upn -AccessRights SendAs -Confirm:$false -WarningAction SilentlyContinue -ErrorAction Stop | Out-Null
                $saRemovedNames += "$($r.Identity)"
            } catch { $accFailed += "SendAs:$($r.Identity)" }
        }
    } catch {
        $accFailed += "SendAs-lookup ($($_.Exception.Message))"
    }
    # FullAccess on shared mailboxes - resolved from the pre-built index.
    if ($null -eq $FullAccessIndex) {
        $accFailed += 'FullAccess:index-unavailable'
    } else {
        $faTargets = if ($FullAccessIndex.ContainsKey($Upn.ToLower())) { $FullAccessIndex[$Upn.ToLower()] } else { @() }
        foreach ($mbxId in $faTargets) {
            try {
                Remove-MailboxPermission -Identity $mbxId -User $Upn -AccessRights FullAccess -Confirm:$false -WarningAction SilentlyContinue -ErrorAction Stop | Out-Null
                $faRemovedNames += "$mbxId"
            } catch { $accFailed += "FullAccess:$mbxId" }
        }
    }
    $saRemoved = $saRemovedNames.Count
    $faRemoved = $faRemovedNames.Count
    $accStatus = if ($accFailed.Count -gt 0) { if (($saRemoved + $faRemoved) -gt 0) { 'Partial' } else { 'Failed' } } else { 'Success' }
    $accParts  = @()
    $accParts += if ($saRemovedNames.Count -gt 0) { "SendAs removed ($saRemoved): " + ($saRemovedNames -join ', ') } else { "SendAs removed: 0" }
    $accParts += if ($faRemovedNames.Count -gt 0) { "Shared FullAccess removed ($faRemoved): " + ($faRemovedNames -join ', ') } else { "Shared FullAccess removed: 0" }
    if ($accFailed.Count -gt 0) { $accParts += "Failed: $($accFailed -join ', ')" }
    Add-Result $Upn $Display 'Remove Shared Mailbox Access' $accStatus ($accParts -join ' | ') -KeepDetail

    # 8) Mailbox size check
    # $SizeKnown separates "under cap" from "could not determine size". When the
    # size of an existing mailbox is unknown, license removal is skipped (safety).
    $SizeGB    = $null
    $SizeKnown = $false
    if ($ConvertedToShared -or ($MailboxType -and $MailboxType -ne 'Unknown')) {
        try {
            $st     = Get-MailboxStatistics -Identity $Upn -ErrorAction Stop
            $SizeGB = ConvertTo-GB $st.TotalItemSize
            if ($null -ne $SizeGB) { $SizeKnown = $true }
        } catch {
            Write-Log "Mailbox size lookup failed for $Upn - license removal will be skipped: $($_.Exception.Message)" 'WARN' -NoConsole
        }
    }
    if ($SizeKnown) {
        $overCap = $SizeGB -gt $SizeCapGB
        Add-Result $Upn $Display 'Mailbox Size Check' 'Info' ("$SizeGB GB" + $(if ($overCap) { " (OVER $SizeCapGB GB)" } else { " (under $SizeCapGB GB)" }))
    } else {
        $overCap = $false
        if ($null -eq $MailboxType) {
            Add-Result $Upn $Display 'Mailbox Size Check' 'Info' 'Size unavailable (no mailbox)'
        } else {
            Add-Result $Upn $Display 'Mailbox Size Check' 'Info' 'Size unavailable (lookup failed on an existing mailbox)'
        }
    }

    # 9) Remove license - only when it is safe
    # Safe cases: no mailbox at all, OR a confirmed Shared mailbox whose size is
    # known and <= cap. Every other case retains the license and is flagged.
    # $removeLicense / $licenseSkipReason decide which path is taken below.
    $removeLicense     = $false
    $licenseSkipReason = $null
    if ($null -eq $MailboxType) {
        $removeLicense = $true                      # no mailbox - removal cannot orphan data
    } elseif (-not $ConvertedToShared) {
        $licenseSkipReason = "Mailbox conversion to Shared not confirmed - license RETAINED for safety"
    } elseif (-not $SizeKnown) {
        $licenseSkipReason = "Mailbox size could not be determined - license RETAINED for safety"
    } elseif ($overCap) {
        $licenseSkipReason = "Mailbox $SizeGB GB > $SizeCapGB GB - license RETAINED (required for shared mailbox over $SizeCapGB GB)"
    } else {
        $removeLicense = $true                       # confirmed shared, size known and under cap
    }

    if (-not $removeLicense) {
        Add-Result $Upn $Display 'Remove License' 'Skipped' $licenseSkipReason
    } else {
        $lic = @($u.AssignedLicenses)
        if (-not $lic -or $lic.Count -eq 0) {
            Add-Result $Upn $Display 'Remove License' 'Skipped' 'No licenses assigned'
        } else {
            $licRemoved = @(); $licFailed = @()
            foreach ($l in $lic) {
                $sku  = $l.SkuId
                $name = if ($SkuMap.ContainsKey($sku)) { $SkuMap[$sku] } else { "$sku" }
                try {
                    Set-MgUserLicense -UserId $Id -RemoveLicenses @($sku) -AddLicenses @() -ErrorAction Stop | Out-Null
                    $licRemoved += $name
                } catch { $licFailed += "$name (likely group-assigned)" }
            }
            if ($licFailed.Count -eq 0) {
                Add-Result $Upn $Display 'Remove License' 'Success' ("Removed: " + ($licRemoved -join ', '))
            } elseif ($licRemoved.Count -gt 0) {
                Add-Result $Upn $Display 'Remove License' 'Partial' ("Removed: " + ($licRemoved -join ', ') + " | Failed: " + ($licFailed -join '; '))
            } else {
                Add-Result $Upn $Display 'Remove License' 'Failed' ("Failed: " + ($licFailed -join '; '))
            }
        }
    }
}

Write-Progress -Activity "Offboarding" -Completed

# ---------------------------------------------------------------------------
# EXPORT RESULTS
# ---------------------------------------------------------------------------
try {
    $Results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Log "Results CSV exported: $OutputPath" 'INFO'
} catch {
    Write-Log "Failed to export results CSV: $($_.Exception.Message)" 'ERROR'
}

# ---------------------------------------------------------------------------
# DISCONNECT
# ---------------------------------------------------------------------------
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Disconnect-MgGraph -ErrorAction SilentlyContinue
Write-Log "Sessions disconnected." 'INFO'

# ---------------------------------------------------------------------------
# SUMMARY
# ---------------------------------------------------------------------------
$ScriptStart.Stop()
$Elapsed = $ScriptStart.Elapsed
$RunTime = "{0:D2}h {1:D2}m {2:D2}s {3:D3}ms" -f `
    $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds, $Elapsed.Milliseconds

$FailCount    = ($Results | Where-Object { $_.Status -eq 'Failed'  }).Count
$PartialCount = ($Results | Where-Object { $_.Status -eq 'Partial' }).Count
$Retained     = @($Results | Where-Object { $_.Action -eq 'Remove License' -and $_.Status -eq 'Skipped' -and $_.Detail -like '*RETAINED*' })

Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |                  RUN SUMMARY                    |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Tenant Code         : {0,-27}|" -f $TenantCode)     -ForegroundColor White
Write-Host ("  | Accounts Submitted  : {0,-27}|" -f $Total)          -ForegroundColor White
Write-Host ("  | Processed           : {0,-27}|" -f $ProcessedCount) -ForegroundColor White
Write-Host ("  | Skipped (prompt)    : {0,-27}|" -f $SkippedCount)   -ForegroundColor White
Write-Host ("  | Not Found           : {0,-27}|" -f $NotFoundCount)  -ForegroundColor White
Write-Host ("  | Action Rows Logged  : {0,-27}|" -f $Results.Count)  -ForegroundColor White
Write-Host ("  | Failed Actions      : {0,-27}|" -f $FailCount)      -ForegroundColor White
Write-Host ("  | Partial Actions     : {0,-27}|" -f $PartialCount)   -ForegroundColor White
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Total Run Time      : {0,-27}|" -f $RunTime)        -ForegroundColor Yellow
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan

Write-Log ("Run summary - Submitted: $Total | Processed: $ProcessedCount | Skipped: $SkippedCount | NotFound: $NotFoundCount | Failed: $FailCount | Partial: $PartialCount | Retained-license: $($Retained.Count)") 'INFO' -NoConsole

if ($Retained.Count -gt 0) {
    Write-Host ""
    Write-Host "  [!] LICENSE RETAINED (review these):" -ForegroundColor Yellow
    $Retained | ForEach-Object { Write-Host "      $($_.UserPrincipalName) - $($_.Detail)" -ForegroundColor Yellow }
}

Write-Host ""
Write-Host "  Run log : $LogPath" -ForegroundColor DarkGray
Write-Host "  Results : $OutputPath" -ForegroundColor DarkGray
Write-Host ""
Read-Host "  Press Enter to exit"
