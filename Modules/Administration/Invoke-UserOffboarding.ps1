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
        6. Convert the mailbox to a Shared Mailbox
        7. Remove group memberships (DL, Mail-Enabled Security, M365, Security)
           and the user's access to shared mailboxes (FullAccess + SendAs)
        8. Check mailbox size:
              > 50 GB  -> license removal is SKIPPED and flagged (an unlicensed
                          shared mailbox may not exceed 50 GB)
              <= 50 GB -> remove all assigned licenses (one SKU at a time, so a
                          group-assigned SKU that can't be removed is logged as
                          failed without blocking the direct ones)

    Connects Microsoft Graph FIRST, then Exchange Online, to avoid the MSAL
    "WithLogging" assembly conflict.

    A per-user / per-action results CSV is written to C:\MSP-M365-Utility\ and
    a grouped summary is printed at the end (including the >50 GB retained-
    license flags).

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
#>

# ----------------------------------------------------------
#  CONFIGURATION
# ----------------------------------------------------------
$ErrorActionPreference = 'Stop'
$Host.UI.RawUI.WindowTitle = "M365 User Offboarding"
$SizeCapGB = 50   # unlicensed shared mailbox limit

# ----------------------------------------------------------
#  BANNER
# ----------------------------------------------------------
Clear-Host
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |            M365 User Offboarding Script          |" -ForegroundColor Cyan
Write-Host "  |        Microsoft Graph  *  Exchange Online       |" -ForegroundColor Cyan
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
$OutputFile = "UserOffboarding_${TenantCode}_${Timestamp}.csv"
$OutputPath = Join-Path $OutputRoot $OutputFile

# ----------------------------------------------------------
#  PASTE-LIST INPUT
# ----------------------------------------------------------
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
    Read-Host "`n  Press Enter to exit"; exit 0
}

# ----------------------------------------------------------
#  PREVIEW + STRONG CONFIRMATION (type the count)
# ----------------------------------------------------------
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Red
Write-Host "  |  WARNING - THIS PERMANENTLY OFFBOARDS ACCOUNTS   |" -ForegroundColor Red
Write-Host "  +--------------------------------------------------+" -ForegroundColor Red
Write-Host "  Each account below will be:" -ForegroundColor Yellow
Write-Host "    - Disabled, manager removed, sessions/tokens revoked" -ForegroundColor DarkGray
Write-Host "    - All MFA methods removed" -ForegroundColor DarkGray
Write-Host "    - Converted to a Shared Mailbox" -ForegroundColor DarkGray
Write-Host "    - Removed from all groups and shared-mailbox access" -ForegroundColor DarkGray
Write-Host "    - License removed (unless mailbox > $SizeCapGB GB)" -ForegroundColor DarkGray
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
    Read-Host "`n  Press Enter to exit"; exit 0
}

$ScriptStart = [System.Diagnostics.Stopwatch]::StartNew()

# ----------------------------------------------------------
#  MODULE CHECK + CONNECT (Graph first, then EXO)
# ----------------------------------------------------------
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
        Write-Host "  [!] Module '$Mod' not found. Installing..." -ForegroundColor Yellow
        Install-Module -Name $Mod -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module -Name $Mod -ErrorAction Stop
    Write-Host "  [OK] $Mod loaded." -ForegroundColor Green
}

Write-Host ""
Write-Host "  [2/3] Connecting (Graph first, then Exchange Online)..." -ForegroundColor Cyan
try {
    Connect-MgGraph -Scopes "User.ReadWrite.All","UserAuthenticationMethod.ReadWrite.All","Group.ReadWrite.All","GroupMember.ReadWrite.All","Directory.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "  [OK] Microsoft Graph connected." -ForegroundColor Green
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "  [OK] Exchange Online connected." -ForegroundColor Green
} catch {
    Write-Host "  [ERROR] Failed to connect: $_" -ForegroundColor Red
    Read-Host "`n  Press Enter to exit"; exit 1
}

# Build a SkuId -> friendly name map once (no extra module needed)
$SkuMap = @{}
try {
    $skuResp = Invoke-MgGraphRequest -Method GET -Uri 'v1.0/subscribedSkus' -ErrorAction Stop
    foreach ($s in $skuResp.value) { $SkuMap[$s.skuId] = $s.skuPartNumber }
} catch { }

# ----------------------------------------------------------
#  RESULT LOGGING HELPERS
# ----------------------------------------------------------
$Results = [System.Collections.Generic.List[PSCustomObject]]::new()
function Add-Result {
    param([string]$Upn, [string]$Display, [string]$Action, [string]$Status, [string]$Detail)
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
}

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

# ----------------------------------------------------------
#  PROCESS EACH USER
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [3/3] Offboarding users..." -ForegroundColor Cyan

$Total   = $Users.Count
$Counter = 0

foreach ($Upn in $Users) {
    $Counter++
    $PercentComplete = [math]::Round(($Counter / $Total) * 100, 1)
    Write-Progress -Activity "Offboarding -- Tenant: $TenantCode" `
                   -Status "$Counter / $Total  ($PercentComplete%)  |  $Upn" `
                   -PercentComplete $PercentComplete

    Write-Host ""
    Write-Host "  === $Upn ===" -ForegroundColor White

    # Resolve user
    $u = $null
    try {
        $u = Get-MgUser -UserId $Upn -Property "Id,DisplayName,UserPrincipalName,AccountEnabled,AssignedLicenses" -ErrorAction Stop
    } catch {
        Add-Result $Upn '' 'Resolve User' 'Failed' "User not found in Graph: $($_.Exception.Message)"
        continue
    }
    $Id      = $u.Id
    $Display = $u.DisplayName

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
    $hasMailbox = $false
    try {
        Set-Mailbox -Identity $Upn -Type Shared -ErrorAction Stop
        $hasMailbox = $true
        Add-Result $Upn $Display 'Convert to Shared' 'Success' 'Mailbox type = Shared'
    } catch {
        Add-Result $Upn $Display 'Convert to Shared' 'Failed' $_.Exception.Message
    }

    # 7a) Remove group memberships
    $grpRemoved = 0; $grpFailed = @(); $grpDynamic = @()
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

            try {
                if ($useExo) {
                    $ident = if ($gMail) { $gMail } else { $gName }
                    Remove-DistributionGroupMember -Identity $ident -Member $Upn -Confirm:$false -ErrorAction Stop
                } else {
                    Remove-MgGroupMemberByRef -GroupId $mo.Id -DirectoryObjectId $Id -ErrorAction Stop
                }
                $grpRemoved++
            } catch { $grpFailed += "$gName ($($_.Exception.Message))" }
        }
        $detailParts = @("Removed $grpRemoved")
        if ($grpDynamic.Count -gt 0) { $detailParts += "Dynamic (skipped): $($grpDynamic -join ', ')" }
        if ($grpFailed.Count  -gt 0) { $detailParts += "Failed: $($grpFailed -join '; ')" }
        $grpStatus = if ($grpFailed.Count -gt 0) { if ($grpRemoved -gt 0) { 'Partial' } else { 'Failed' } } else { 'Success' }
        Add-Result $Upn $Display 'Remove Group Memberships' $grpStatus ($detailParts -join ' | ')
    } catch {
        Add-Result $Upn $Display 'Remove Group Memberships' 'Failed' $_.Exception.Message
    }

    # 7b) Remove shared-mailbox access (SendAs org-wide; FullAccess on shared mailboxes)
    $saRemoved = 0; $faRemoved = 0; $accFailed = @()
    # SendAs
    try {
        $sendAs = @(Get-RecipientPermission -Trustee $Upn -ResultSize Unlimited -ErrorAction Stop |
                    Where-Object { $_.AccessRights -contains 'SendAs' })
        foreach ($r in $sendAs) {
            try {
                Remove-RecipientPermission -Identity $r.Identity -Trustee $Upn -AccessRights SendAs -Confirm:$false -WarningAction SilentlyContinue -ErrorAction Stop | Out-Null
                $saRemoved++
            } catch { $accFailed += "SendAs:$($r.Identity)" }
        }
    } catch { }
    # FullAccess on shared mailboxes only
    try {
        $sharedMbx = @(Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop)
        foreach ($sm in $sharedMbx) {
            try {
                $fa = Get-MailboxPermission -Identity $sm.Identity -User $Upn -ErrorAction SilentlyContinue |
                      Where-Object { $_.AccessRights -contains 'FullAccess' -and -not $_.IsInherited }
                if ($fa) {
                    Remove-MailboxPermission -Identity $sm.Identity -User $Upn -AccessRights FullAccess -Confirm:$false -WarningAction SilentlyContinue -ErrorAction Stop | Out-Null
                    $faRemoved++
                }
            } catch { $accFailed += "FullAccess:$($sm.PrimarySmtpAddress)" }
        }
    } catch { }
    $accStatus = if ($accFailed.Count -gt 0) { 'Partial' } else { 'Success' }
    $accDetail = "SendAs removed: $saRemoved | Shared FullAccess removed: $faRemoved"
    if ($accFailed.Count -gt 0) { $accDetail += " | Failed: $($accFailed -join ', ')" }
    Add-Result $Upn $Display 'Remove Shared Mailbox Access' $accStatus $accDetail

    # 8) Mailbox size check
    $SizeGB = $null
    if ($hasMailbox) {
        try {
            $st     = Get-MailboxStatistics -Identity $Upn -ErrorAction Stop
            $SizeGB = ConvertTo-GB $st.TotalItemSize
        } catch { }
    }
    if ($null -ne $SizeGB) {
        $overCap = $SizeGB -gt $SizeCapGB
        Add-Result $Upn $Display 'Mailbox Size Check' 'Info' ("$SizeGB GB" + $(if ($overCap) { " (OVER $SizeCapGB GB)" } else { " (under $SizeCapGB GB)" }))
    } else {
        $overCap = $false
        Add-Result $Upn $Display 'Mailbox Size Check' 'Info' 'Size unavailable (no accessible mailbox)'
    }

    # 9) Remove license (unless over cap)
    if ($overCap) {
        Add-Result $Upn $Display 'Remove License' 'Skipped' "Mailbox $SizeGB GB > $SizeCapGB GB - license RETAINED (required for shared mailbox over $SizeCapGB GB)"
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
Disconnect-MgGraph -ErrorAction SilentlyContinue
Write-Host "  Sessions disconnected." -ForegroundColor DarkGray

# ----------------------------------------------------------
#  SUMMARY
# ----------------------------------------------------------
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
Write-Host ("  | Tenant Code         : {0,-27}|" -f $TenantCode)   -ForegroundColor White
Write-Host ("  | Accounts Processed  : {0,-27}|" -f $Total)        -ForegroundColor White
Write-Host ("  | Action Rows Logged  : {0,-27}|" -f $Results.Count)-ForegroundColor White
Write-Host ("  | Failed Actions      : {0,-27}|" -f $FailCount)    -ForegroundColor White
Write-Host ("  | Partial Actions     : {0,-27}|" -f $PartialCount) -ForegroundColor White
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Total Run Time      : {0,-27}|" -f $RunTime)      -ForegroundColor Yellow
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan

if ($Retained.Count -gt 0) {
    Write-Host ""
    Write-Host "  [!] LICENSE RETAINED (mailbox over $SizeCapGB GB) - review these:" -ForegroundColor Yellow
    $Retained | ForEach-Object { Write-Host "      $($_.UserPrincipalName) - $($_.Detail)" -ForegroundColor Yellow }
}

Write-Host ""
Read-Host "  Press Enter to exit"
