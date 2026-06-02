#Requires -Modules ExchangeOnlineManagement, Microsoft.Graph.Users, Microsoft.Graph.Groups

<#
.SYNOPSIS
    Retrieves all groups and their members from a Microsoft 365 tenant and exports to CSV.

.DESCRIPTION
    Connects to Exchange Online and Microsoft Graph to enumerate all group types:
        - Distribution Groups
        - Mail-Enabled Security Groups
        - Microsoft 365 Groups
        - Security Groups (non-mail-enabled, via Graph)
    For each group, all members are resolved and classified as:
        - Internal User
        - Shared Mailbox
        - External Contact
        - Nested Distribution Group
        - Nested Mail-Enabled Security Group
        - Nested Microsoft 365 Group
        - Nested Security Group
    Results are exported to a flat CSV with one row per group-member combination.

.NOTES
    Required Modules:
        - ExchangeOnlineManagement  : Install-Module ExchangeOnlineManagement
        - Microsoft.Graph.Users     : Install-Module Microsoft.Graph.Users
        - Microsoft.Graph.Groups    : Install-Module Microsoft.Graph.Groups
    Required Permissions:
        - Exchange Online : View-Only Recipients (or higher)
        - Microsoft Graph : User.Read.All, Group.Read.All, GroupMember.Read.All
#>

# ----------------------------------------------------------
#  CONFIGURATION
# ----------------------------------------------------------
$ErrorActionPreference = 'Stop'
$Host.UI.RawUI.WindowTitle = "M365 Group Membership Inventory"

# ----------------------------------------------------------
#  BANNER
# ----------------------------------------------------------
Clear-Host
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |     M365 Group Membership Inventory Script       |" -ForegroundColor Cyan
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
$OutputFile = "GroupMembership_${TenantCode}_${Timestamp}.csv"
$OutputPath = Join-Path $PSScriptRoot $OutputFile

Write-Host ""
Write-Host "  Tenant Code : $TenantCode" -ForegroundColor Green
Write-Host "  Output File : $OutputPath" -ForegroundColor Green
Write-Host ""

# ----------------------------------------------------------
#  START TIMER
# ----------------------------------------------------------
$ScriptStart = [System.Diagnostics.Stopwatch]::StartNew()

# ----------------------------------------------------------
#  MODULE CHECK
# ----------------------------------------------------------
Write-Host "  [1/4] Checking required modules..." -ForegroundColor Cyan

$RequiredModules = @(
    'ExchangeOnlineManagement',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Groups'
)
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

    Connect-MgGraph -Scopes "User.Read.All","Group.Read.All","GroupMember.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "  [OK] Microsoft Graph connected." -ForegroundColor Green
}
catch {
    Write-Host "  [ERROR] Failed to connect: $_" -ForegroundColor Red
    Read-Host "`n  Press Enter to exit"
    exit 1
}

# ----------------------------------------------------------
#  RETRIEVE GROUPS + PRE-FETCH LOOKUP DATA
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [3/4] Retrieving groups and building lookup tables..." -ForegroundColor Cyan

$AllGroups = [System.Collections.Generic.List[PSCustomObject]]::new()

# Distribution Groups
try {
    Write-Host "  Fetching Distribution Groups..." -ForegroundColor DarkCyan
    $DistGroups = Get-DistributionGroup -RecipientTypeDetails MailUniversalDistributionGroup -ResultSize Unlimited -ErrorAction Stop
    foreach ($G in $DistGroups) {
        $AllGroups.Add([PSCustomObject]@{
            GroupName  = $G.DisplayName
            GroupEmail = $G.PrimarySmtpAddress
            GroupType  = "Distribution Group"
            Identity   = $G.Identity
            Source     = "EXO-DG"
        })
    }
    Write-Host "  [OK] Distribution Groups          : $($DistGroups.Count)" -ForegroundColor Green
}
catch {
    Write-Host "  [WARN] Could not retrieve Distribution Groups: $_" -ForegroundColor Yellow
}

# Mail-Enabled Security Groups
try {
    Write-Host "  Fetching Mail-Enabled Security Groups..." -ForegroundColor DarkCyan
    $MailSecGroups = Get-DistributionGroup -RecipientTypeDetails MailUniversalSecurityGroup -ResultSize Unlimited -ErrorAction Stop
    foreach ($G in $MailSecGroups) {
        $AllGroups.Add([PSCustomObject]@{
            GroupName  = $G.DisplayName
            GroupEmail = $G.PrimarySmtpAddress
            GroupType  = "Mail-Enabled Security Group"
            Identity   = $G.Identity
            Source     = "EXO-MESG"
        })
    }
    Write-Host "  [OK] Mail-Enabled Security Groups : $($MailSecGroups.Count)" -ForegroundColor Green
}
catch {
    Write-Host "  [WARN] Could not retrieve Mail-Enabled Security Groups: $_" -ForegroundColor Yellow
}

# Microsoft 365 Groups
try {
    Write-Host "  Fetching Microsoft 365 Groups..." -ForegroundColor DarkCyan
    $UnifiedGroups = Get-UnifiedGroup -ResultSize Unlimited -ErrorAction Stop
    foreach ($G in $UnifiedGroups) {
        $AllGroups.Add([PSCustomObject]@{
            GroupName  = $G.DisplayName
            GroupEmail = $G.PrimarySmtpAddress
            GroupType  = "Microsoft 365 Group"
            Identity   = $G.Identity
            Source     = "EXO-M365"
        })
    }
    Write-Host "  [OK] Microsoft 365 Groups         : $($UnifiedGroups.Count)" -ForegroundColor Green
}
catch {
    Write-Host "  [WARN] Could not retrieve Microsoft 365 Groups: $_" -ForegroundColor Yellow
}

# Security Groups (non-mail-enabled) via Graph
try {
    Write-Host "  Fetching Security Groups (non-mail-enabled) via Graph..." -ForegroundColor DarkCyan
    $GraphSecGroups = Get-MgGroup -All `
        -Filter "securityEnabled eq true and mailEnabled eq false" `
        -Property "Id,DisplayName,Mail,GroupTypes" `
        -ErrorAction Stop |
        Where-Object { $_.GroupTypes -notcontains "Unified" }

    foreach ($G in $GraphSecGroups) {
        $AllGroups.Add([PSCustomObject]@{
            GroupName  = $G.DisplayName
            GroupEmail = $G.Mail
            GroupType  = "Security Group"
            Identity   = $G.Id
            Source     = "Graph-SG"
        })
    }
    Write-Host "  [OK] Security Groups (Graph)      : $($GraphSecGroups.Count)" -ForegroundColor Green
}
catch {
    Write-Host "  [WARN] Could not retrieve Security Groups from Graph: $_" -ForegroundColor Yellow
}

$TotalGroups = $AllGroups.Count
Write-Host ""
Write-Host "  Total Groups Found : $TotalGroups" -ForegroundColor Green

# ----------------------------------------------------------
#  BUILD MEMBER LOOKUP TABLES
# ----------------------------------------------------------
Write-Host ""
Write-Host "  Building member lookup tables..." -ForegroundColor DarkCyan

# Shared Mailboxes - keyed by PrimarySmtpAddress
$SharedMailboxLookup = @{}
try {
    Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop |
        ForEach-Object {
            if ($_.PrimarySmtpAddress) {
                $SharedMailboxLookup[$_.PrimarySmtpAddress.ToLower()] = $true
            }
        }
    Write-Host "  [OK] Shared Mailbox lookup built     ($($SharedMailboxLookup.Count) entries)." -ForegroundColor Green
}
catch {
    Write-Host "  [WARN] Could not build Shared Mailbox lookup: $_" -ForegroundColor Yellow
}

# Mail Contacts - keyed by PrimarySmtpAddress
$MailContactLookup = @{}
try {
    Get-MailContact -ResultSize Unlimited -ErrorAction Stop |
        ForEach-Object {
            if ($_.PrimarySmtpAddress) {
                $MailContactLookup[$_.PrimarySmtpAddress.ToLower()] = $true
            }
        }
    Write-Host "  [OK] Mail Contact lookup built       ($($MailContactLookup.Count) entries)." -ForegroundColor Green
}
catch {
    Write-Host "  [WARN] Could not build Mail Contact lookup: $_" -ForegroundColor Yellow
}

# Distribution Groups - keyed by PrimarySmtpAddress
# Used as a fallback to positively identify nested DLs even when
# RecipientTypeDetails is missing or ambiguous (e.g. dynamic members)
$DistGroupLookup = @{}
try {
    Get-DistributionGroup -RecipientTypeDetails MailUniversalDistributionGroup -ResultSize Unlimited -ErrorAction Stop |
        ForEach-Object {
            if ($_.PrimarySmtpAddress) {
                $DistGroupLookup[$_.PrimarySmtpAddress.ToLower()] = $true
            }
        }
    Write-Host "  [OK] Distribution Group lookup built ($($DistGroupLookup.Count) entries)." -ForegroundColor Green
}
catch {
    Write-Host "  [WARN] Could not build Distribution Group lookup: $_" -ForegroundColor Yellow
}

# Mail-Enabled Security Groups - keyed by PrimarySmtpAddress
$MailSecGroupLookup = @{}
try {
    Get-DistributionGroup -RecipientTypeDetails MailUniversalSecurityGroup -ResultSize Unlimited -ErrorAction Stop |
        ForEach-Object {
            if ($_.PrimarySmtpAddress) {
                $MailSecGroupLookup[$_.PrimarySmtpAddress.ToLower()] = $true
            }
        }
    Write-Host "  [OK] Mail-Enabled Sec Group lookup   ($($MailSecGroupLookup.Count) entries)." -ForegroundColor Green
}
catch {
    Write-Host "  [WARN] Could not build Mail-Enabled Security Group lookup: $_" -ForegroundColor Yellow
}

# ----------------------------------------------------------
#  HELPER: RESOLVE GRAPH GROUP TYPE
# ----------------------------------------------------------
# When a Graph security group member is itself a group, fetch its
# properties and return the correct specific nested label.
function Resolve-GraphGroupType {
    param ([string]$GroupId)

    try {
        $G = Get-MgGroup -GroupId $GroupId `
            -Property "GroupTypes,SecurityEnabled,MailEnabled" `
            -ErrorAction Stop

        if ($G.GroupTypes -contains "Unified")             { return "Nested Microsoft 365 Group"         }
        if ($G.MailEnabled  -and $G.SecurityEnabled)       { return "Nested Mail-Enabled Security Group" }
        if ($G.MailEnabled  -and -not $G.SecurityEnabled)  { return "Nested Distribution Group"          }
        if ($G.SecurityEnabled -and -not $G.MailEnabled)   { return "Nested Security Group"              }
    }
    catch {
        # Lookup failed - fall through to generic label
    }

    return "Nested Group"
}

# ----------------------------------------------------------
#  HELPER: RESOLVE MEMBER TYPE
# ----------------------------------------------------------
function Resolve-MemberType {
    param (
        [string]$RecipientTypeDetails,
        [string]$EmailAddress,
        [string]$ODataType,       # Populated for Graph-sourced members
        [string]$GraphGroupType   # Populated when Graph member is a group
    )

    # --- Graph path ---
    if ($ODataType) {
        if ($ODataType -like "*group*") {
            # Return the specific group type resolved by the caller
            if ($GraphGroupType) { return $GraphGroupType }
            return "Nested Group"
        }
        if ($ODataType -like "*user*") { return "Internal User" }
        return "Unknown"
    }

    # --- EXO path: use RecipientTypeDetails directly ---
    switch ($RecipientTypeDetails) {
        "UserMailbox"                    { return "Internal User"                       }
        "GuestMailUser"                  { return "External Contact"                    }
        "MailContact"                    { return "External Contact"                    }
        "SharedMailbox"                  { return "Shared Mailbox"                      }
        "MailUniversalDistributionGroup" { return "Nested Distribution Group"           }
        "MailUniversalSecurityGroup"     { return "Nested Mail-Enabled Security Group"  }
        "GroupMailbox"                   { return "Nested Microsoft 365 Group"          }
        default {
            # Fallback: check pre-fetched lookup tables
            if ($EmailAddress) {
                $Key = $EmailAddress.ToLower()
                if ($DistGroupLookup.ContainsKey($Key))     { return "Nested Distribution Group"          }
                if ($MailSecGroupLookup.ContainsKey($Key))  { return "Nested Mail-Enabled Security Group" }
                if ($SharedMailboxLookup.ContainsKey($Key)) { return "Shared Mailbox"                     }
                if ($MailContactLookup.ContainsKey($Key))   { return "External Contact"                   }
            }
            if ($RecipientTypeDetails) { return $RecipientTypeDetails }
            return "Unknown"
        }
    }
}

# ----------------------------------------------------------
#  PROCESS GROUPS + MEMBERS WITH PROGRESS BAR
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [4/4] Processing groups and resolving members..." -ForegroundColor Cyan
Write-Host ""

$Results      = [System.Collections.Generic.List[PSCustomObject]]::new()
$GroupCounter = 0

foreach ($Group in $AllGroups) {
    $GroupCounter++

    # Progress Bar
    $PercentComplete = [math]::Round(($GroupCounter / $TotalGroups) * 100, 1)
    $BarWidth        = 40
    $FilledBars      = [math]::Floor($PercentComplete / 100 * $BarWidth)
    $EmptyBars       = $BarWidth - $FilledBars
    $ProgressBar     = "[" + ("#" * $FilledBars) + ("-" * $EmptyBars) + "]"

    Write-Progress `
        -Activity        "Processing Groups -- Tenant: $TenantCode" `
        -Status          "$ProgressBar  $GroupCounter / $TotalGroups  ($PercentComplete%)  |  $($Group.GroupName)" `
        -PercentComplete  $PercentComplete `
        -CurrentOperation "Source: $($Group.Source)  |  Type: $($Group.GroupType)"

    $Members = @()

    switch ($Group.Source) {

        # -------------------------------------------------------
        # Distribution Groups and Mail-Enabled Security Groups
        # -------------------------------------------------------
        { $_ -in "EXO-DG", "EXO-MESG" } {
            try {
                $Members = Get-DistributionGroupMember `
                    -Identity   $Group.Identity `
                    -ResultSize Unlimited `
                    -ErrorAction Stop
            }
            catch {
                Write-Host "  [WARN] Could not get members for '$($Group.GroupName)': $_" -ForegroundColor Yellow
            }

            if ($Members.Count -eq 0) {
                $Results.Add([PSCustomObject]@{
                    GroupName          = $Group.GroupName
                    GroupEmailAddress  = $Group.GroupEmail
                    GroupType          = $Group.GroupType
                    MemberDisplayName  = "(No Members)"
                    MemberEmailAddress = $null
                    MemberType         = $null
                })
            }
            else {
                foreach ($Member in $Members) {
                    $MemberType = Resolve-MemberType `
                        -RecipientTypeDetails $Member.RecipientTypeDetails `
                        -EmailAddress         $Member.PrimarySmtpAddress

                    $Results.Add([PSCustomObject]@{
                        GroupName          = $Group.GroupName
                        GroupEmailAddress  = $Group.GroupEmail
                        GroupType          = $Group.GroupType
                        MemberDisplayName  = $Member.DisplayName
                        MemberEmailAddress = $Member.PrimarySmtpAddress
                        MemberType         = $MemberType
                    })
                }
            }
        }

        # -------------------------------------------------------
        # Microsoft 365 Groups
        # -------------------------------------------------------
        "EXO-M365" {
            try {
                $Members = Get-UnifiedGroupLinks `
                    -Identity   $Group.Identity `
                    -LinkType   Members `
                    -ResultSize Unlimited `
                    -ErrorAction Stop
            }
            catch {
                Write-Host "  [WARN] Could not get members for '$($Group.GroupName)': $_" -ForegroundColor Yellow
            }

            if ($Members.Count -eq 0) {
                $Results.Add([PSCustomObject]@{
                    GroupName          = $Group.GroupName
                    GroupEmailAddress  = $Group.GroupEmail
                    GroupType          = $Group.GroupType
                    MemberDisplayName  = "(No Members)"
                    MemberEmailAddress = $null
                    MemberType         = $null
                })
            }
            else {
                foreach ($Member in $Members) {
                    $MemberType = Resolve-MemberType `
                        -RecipientTypeDetails $Member.RecipientTypeDetails `
                        -EmailAddress         $Member.PrimarySmtpAddress

                    $Results.Add([PSCustomObject]@{
                        GroupName          = $Group.GroupName
                        GroupEmailAddress  = $Group.GroupEmail
                        GroupType          = $Group.GroupType
                        MemberDisplayName  = $Member.DisplayName
                        MemberEmailAddress = $Member.PrimarySmtpAddress
                        MemberType         = $MemberType
                    })
                }
            }
        }

        # -------------------------------------------------------
        # Security Groups (non-mail-enabled) via Graph
        # -------------------------------------------------------
        "Graph-SG" {
            try {
                $Members = Get-MgGroupMember -GroupId $Group.Identity -All -ErrorAction Stop
            }
            catch {
                Write-Host "  [WARN] Could not get members for '$($Group.GroupName)': $_" -ForegroundColor Yellow
            }

            if ($Members.Count -eq 0) {
                $Results.Add([PSCustomObject]@{
                    GroupName          = $Group.GroupName
                    GroupEmailAddress  = $Group.GroupEmail
                    GroupType          = $Group.GroupType
                    MemberDisplayName  = "(No Members)"
                    MemberEmailAddress = $null
                    MemberType         = $null
                })
            }
            else {
                foreach ($Member in $Members) {
                    $ODataType   = $Member.AdditionalProperties['@odata.type']
                    $DisplayName = $Member.AdditionalProperties['displayName']
                    $Email       = $Member.AdditionalProperties['mail']
                    $UserType    = $Member.AdditionalProperties['userType']

                    # When the member is a group, look up its exact type
                    $GraphGroupType = $null
                    if ($ODataType -like "*group*") {
                        $GraphGroupType = Resolve-GraphGroupType -GroupId $Member.Id
                    }

                    $MemberType = Resolve-MemberType `
                        -ODataType      $ODataType `
                        -GraphGroupType $GraphGroupType

                    # Guests reported as users are External Contacts
                    if ($MemberType -eq "Internal User" -and $UserType -eq "Guest") {
                        $MemberType = "External Contact"
                    }

                    $Results.Add([PSCustomObject]@{
                        GroupName          = $Group.GroupName
                        GroupEmailAddress  = $Group.GroupEmail
                        GroupType          = $Group.GroupType
                        MemberDisplayName  = $DisplayName
                        MemberEmailAddress = $Email
                        MemberType         = $MemberType
                    })
                }
            }
        }
    }
}

Write-Progress -Activity "Processing Groups" -Completed

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

$DGCount         = ($AllGroups | Where-Object { $_.GroupType -eq "Distribution Group"          }).Count
$MESGCount       = ($AllGroups | Where-Object { $_.GroupType -eq "Mail-Enabled Security Group" }).Count
$M365Count       = ($AllGroups | Where-Object { $_.GroupType -eq "Microsoft 365 Group"         }).Count
$SGCount         = ($AllGroups | Where-Object { $_.GroupType -eq "Security Group"              }).Count

$InternalCount   = ($Results   | Where-Object { $_.MemberType -eq "Internal User"                      }).Count
$SharedCount     = ($Results   | Where-Object { $_.MemberType -eq "Shared Mailbox"                     }).Count
$ExternalCount   = ($Results   | Where-Object { $_.MemberType -eq "External Contact"                   }).Count
$NestedDLCount   = ($Results   | Where-Object { $_.MemberType -eq "Nested Distribution Group"          }).Count
$NestedMESGCount = ($Results   | Where-Object { $_.MemberType -eq "Nested Mail-Enabled Security Group" }).Count
$NestedM365Count = ($Results   | Where-Object { $_.MemberType -eq "Nested Microsoft 365 Group"         }).Count
$NestedSGCount   = ($Results   | Where-Object { $_.MemberType -eq "Nested Security Group"              }).Count
$EmptyCount      = ($Results   | Where-Object { $_.MemberDisplayName -eq "(No Members)"                }).Count
$TotalRows       = $Results.Count

$RunTime = "{0:D2}h {1:D2}m {2:D2}s {3:D3}ms" -f `
    $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds, $Elapsed.Milliseconds

Write-Host ""
Write-Host "  +------------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |                    RUN SUMMARY                      |" -ForegroundColor Cyan
Write-Host "  +------------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Tenant Code                       : {0,-15}|" -f $TenantCode)      -ForegroundColor White
Write-Host "  |                                                      |" -ForegroundColor Cyan
Write-Host ("  | Total Groups                      : {0,-15}|" -f $TotalGroups)     -ForegroundColor White
Write-Host ("  |   Distribution Groups             : {0,-15}|" -f $DGCount)         -ForegroundColor White
Write-Host ("  |   Mail-Enabled Security Groups    : {0,-15}|" -f $MESGCount)       -ForegroundColor White
Write-Host ("  |   Microsoft 365 Groups            : {0,-15}|" -f $M365Count)       -ForegroundColor White
Write-Host ("  |   Security Groups                 : {0,-15}|" -f $SGCount)         -ForegroundColor White
Write-Host "  |                                                      |" -ForegroundColor Cyan
Write-Host ("  | Total CSV Rows                    : {0,-15}|" -f $TotalRows)       -ForegroundColor White
Write-Host ("  |   Internal Users                  : {0,-15}|" -f $InternalCount)   -ForegroundColor White
Write-Host ("  |   Shared Mailboxes                : {0,-15}|" -f $SharedCount)     -ForegroundColor White
Write-Host ("  |   External Contacts               : {0,-15}|" -f $ExternalCount)   -ForegroundColor White
Write-Host ("  |   Nested Distribution Groups      : {0,-15}|" -f $NestedDLCount)   -ForegroundColor White
Write-Host ("  |   Nested Mail-Enabled Sec Groups  : {0,-15}|" -f $NestedMESGCount) -ForegroundColor White
Write-Host ("  |   Nested Microsoft 365 Groups     : {0,-15}|" -f $NestedM365Count) -ForegroundColor White
Write-Host ("  |   Nested Security Groups          : {0,-15}|" -f $NestedSGCount)   -ForegroundColor White
Write-Host ("  |   Empty Groups                    : {0,-15}|" -f $EmptyCount)      -ForegroundColor White
Write-Host "  +------------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Total Run Time                    : {0,-15}|" -f $RunTime)         -ForegroundColor Yellow
Write-Host "  +------------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""
Read-Host "  Press Enter to exit"
