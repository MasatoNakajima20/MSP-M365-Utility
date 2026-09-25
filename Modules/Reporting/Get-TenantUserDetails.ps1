#Requires -Modules Microsoft.Graph.Users

<#
.SYNOPSIS
    Retrieves user details from a Microsoft 365 tenant and exports to CSV.

.DESCRIPTION
    Connects to Microsoft Graph to gather user information including
    FirstName, LastName, PrimarySmtpAddress, AccountStatus (Enabled/Disabled),
    Title, Department, Manager, City, State, Country, PhoneNumber (business),
    MobileNumber, Licenses (friendly product names), and Roles (directory roles).
    Displays a real-time progress bar and exports results to CSV.

    Scope is selectable at runtime:
        Users (default) - userType eq 'Member' (excludes guests)
        All             - every user object, including guests

.NOTES
    Required Modules:
        - Microsoft.Graph.Users : Install-Module Microsoft.Graph.Users
    Required Permissions:
        - Microsoft Graph : User.Read.All
                           OrgContact.Read.All (for Manager lookup)
                           Directory.Read.All  (for License and Role lookup)
#>

# ----------------------------------------------------------
#  CONFIGURATION
# ----------------------------------------------------------
$ErrorActionPreference = 'Stop'
$Host.UI.RawUI.WindowTitle = "M365 User Details Inventory"

# ----------------------------------------------------------
#  SKU FRIENDLY-NAME LOOKUP
# ----------------------------------------------------------
#  Maps the raw SkuPartNumber that Microsoft returns to a human-friendly
#  product name. Any SKU not listed here falls back to its raw SkuPartNumber
#  so nothing is ever blank. Add new SKUs here as Microsoft introduces them.
$SkuFriendlyNames = @{
    'ENTERPRISEPACK'                      = 'Office 365 E3'
    'ENTERPRISEPREMIUM'                   = 'Office 365 E5'
    'ENTERPRISEPREMIUM_NOPSTNCONF'        = 'Office 365 E5 (without Audio Conferencing)'
    'STANDARDPACK'                        = 'Office 365 E1'
    'DESKLESSPACK'                        = 'Office 365 F3'
    'SPE_E3'                              = 'Microsoft 365 E3'
    'SPE_E5'                              = 'Microsoft 365 E5'
    'SPE_F1'                              = 'Microsoft 365 F3'
    'SPE_F5_SECCOMP'                      = 'Microsoft 365 F5 Security + Compliance'
    'SPB'                                = 'Microsoft 365 Business Premium'
    'O365_BUSINESS_PREMIUM'               = 'Microsoft 365 Business Standard'
    'O365_BUSINESS_ESSENTIALS'            = 'Microsoft 365 Business Basic'
    'O365_BUSINESS'                       = 'Microsoft 365 Apps for Business'
    'OFFICESUBSCRIPTION'                  = 'Microsoft 365 Apps for Enterprise'
    'EXCHANGESTANDARD'                    = 'Exchange Online (Plan 1)'
    'EXCHANGEENTERPRISE'                  = 'Exchange Online (Plan 2)'
    'EXCHANGEDESKLESS'                    = 'Exchange Online Kiosk'
    'EXCHANGE_S_ESSENTIALS'               = 'Exchange Online Essentials'
    'FLOW_FREE'                           = 'Microsoft Power Automate Free'
    'POWER_BI_STANDARD'                   = 'Power BI (free)'
    'POWER_BI_PRO'                        = 'Power BI Pro'
    'POWERAPPS_PER_USER'                  = 'Power Apps Premium (per user)'
    'PROJECT_P1'                          = 'Project Plan 1'
    'PROJECTPROFESSIONAL'                 = 'Project Plan 3'
    'PROJECTPREMIUM'                      = 'Project Plan 5'
    'VISIO_PLAN1_DEPT'                    = 'Visio Plan 1'
    'VISIOCLIENT'                         = 'Visio Plan 2'
    'EMS'                                = 'Enterprise Mobility + Security E3'
    'EMSPREMIUM'                          = 'Enterprise Mobility + Security E5'
    'AAD_PREMIUM'                         = 'Entra ID P1'
    'AAD_PREMIUM_P2'                      = 'Entra ID P2'
    'RIGHTSMANAGEMENT'                    = 'Azure Information Protection Plan 1'
    'ATP_ENTERPRISE'                      = 'Microsoft Defender for Office 365 (Plan 1)'
    'THREAT_INTELLIGENCE'                 = 'Microsoft Defender for Office 365 (Plan 2)'
    'WINDEFATP'                           = 'Microsoft Defender for Endpoint'
    'IDENTITY_THREAT_PROTECTION'          = 'Microsoft 365 E5 Security'
    'INFORMATION_PROTECTION_COMPLIANCE'   = 'Microsoft 365 E5 Compliance'
    'MCOEV'                              = 'Microsoft Teams Phone Standard'
    'MCOMEETADV'                          = 'Microsoft 365 Audio Conferencing'
    'PHONESYSTEM_VIRTUALUSER'             = 'Microsoft Teams Phone Resource Account'
    'MCOPSTN1'                            = 'Domestic Calling Plan'
    'MCOPSTN2'                            = 'International Calling Plan'
    'MCOPSTNC'                            = 'Communications Credits'
    'TEAMS_EXPLORATORY'                   = 'Microsoft Teams Exploratory'
    'MS_TEAMS_IW'                         = 'Microsoft Teams Trial'
    'STREAM'                             = 'Microsoft Stream'
    'WIN10_PRO_ENT_SUB'                   = 'Windows 10/11 Enterprise E3'
    'WIN10_VDA_E3'                        = 'Windows 10/11 Enterprise E3'
    'WIN10_VDA_E5'                        = 'Windows 10/11 Enterprise E5'
    'Microsoft_365_Copilot'               = 'Microsoft 365 Copilot'
}

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

# ----------------------------------------------------------
#  INPUT - SCOPE (All users, or Members/Users only)
# ----------------------------------------------------------
#   Users (default) : userType eq 'Member' - excludes guests
#   All             : every user object, including guests
$ScopeAnswer = (Read-Host "  Report scope - [U]sers only (Members) / [A]ll users  [U]").Trim().ToUpper()
$Scope       = if ($ScopeAnswer -eq 'A') { 'All' } else { 'Users' }

# Output path (C:\MSP-M365-Utility\)
$OutputRoot = 'C:\MSP-M365-Utility'
if (-not (Test-Path $OutputRoot)) {
    try {
        New-Item -ItemType Directory -Path $OutputRoot -Force -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Host "  [ERROR] Could not create output folder '$OutputRoot': $_" -ForegroundColor Red
        Read-Host "`n  Press Enter to exit"
        exit 1
    }
}
$Timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputFile = "UserDetails_${TenantCode}_${Timestamp}.csv"
$OutputPath = Join-Path $OutputRoot $OutputFile

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
    Connect-MgGraph -Scopes "User.Read.All","OrgContact.Read.All","Directory.Read.All" -NoWelcome -ErrorAction Stop
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
Write-Host "  [3/4] Retrieving users from Microsoft Graph (scope: $Scope)..." -ForegroundColor Cyan

try {
    $GetUserParams = @{
        All         = $true
        Property    = "GivenName,Surname,Mail,JobTitle,Department,UserPrincipalName,Id,City,State,Country,BusinessPhones,MobilePhone,UserType,AccountEnabled"
        ErrorAction = 'Stop'
    }
    if ($Scope -eq 'Users') {
        $GetUserParams['Filter'] = "userType eq 'Member'"
    }

    $Users = Get-MgUser @GetUserParams

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

    $BusinessPhone = if ($User.BusinessPhones -and $User.BusinessPhones.Count -gt 0) { $User.BusinessPhones[0] } else { $null }
    $AccountStatus = if ($User.AccountEnabled -eq $true) { 'Enabled' } elseif ($User.AccountEnabled -eq $false) { 'Disabled' } else { 'Unknown' }

    # License lookup - resolve each assigned SKU to a friendly product name,
    # falling back to the raw SkuPartNumber for any SKU not in $SkuFriendlyNames.
    $UserLicenses = $null
    try {
        $LicenseDetails = Get-MgUserLicenseDetail -UserId $User.Id -ErrorAction Stop
        $LicenseNames   = foreach ($Lic in $LicenseDetails) {
            if ($SkuFriendlyNames.ContainsKey($Lic.SkuPartNumber)) {
                $SkuFriendlyNames[$Lic.SkuPartNumber]
            }
            else {
                $Lic.SkuPartNumber
            }
        }
        if ($LicenseNames) {
            $UserLicenses = ($LicenseNames | Sort-Object -Unique) -join '; '
        }
    }
    catch {
        $UserLicenses = $null
    }

    # Role lookup - directory roles assigned to the user (e.g. Global
    # Administrator). Get-MgUserMemberOf returns both groups and roles, so
    # keep only the directoryRole objects and read their display names.
    $UserRoles = $null
    try {
        $MemberOf  = Get-MgUserMemberOf -UserId $User.Id -All -ErrorAction Stop
        $RoleNames = foreach ($Obj in $MemberOf) {
            if ($Obj.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.directoryRole') {
                $Obj.AdditionalProperties['displayName']
            }
        }
        if ($RoleNames) {
            $UserRoles = ($RoleNames | Sort-Object -Unique) -join '; '
        }
    }
    catch {
        $UserRoles = $null
    }

    $Results.Add([PSCustomObject]@{
        FirstName          = $User.GivenName
        LastName           = $User.Surname
        PrimarySmtpAddress = if ($User.Mail) { $User.Mail } else { $User.UserPrincipalName }
        AccountStatus      = $AccountStatus
        Title              = $User.JobTitle
        Department         = $User.Department
        Manager            = $ManagerName
        City               = $User.City
        State              = $User.State
        Country            = $User.Country
        PhoneNumber        = $BusinessPhone
        MobileNumber       = $User.MobilePhone
        Licenses           = $UserLicenses
        Roles              = $UserRoles
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
$EnabledCount        = ($Results | Where-Object { $_.AccountStatus -eq 'Enabled'  }).Count
$DisabledCount       = ($Results | Where-Object { $_.AccountStatus -eq 'Disabled' }).Count
$WithLicenseCount    = ($Results | Where-Object { $_.Licenses }).Count
$WithRoleCount       = ($Results | Where-Object { $_.Roles    }).Count

$RunTime = "{0:D2}h {1:D2}m {2:D2}s {3:D3}ms" -f `
    $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds, $Elapsed.Milliseconds

Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |                  RUN SUMMARY                    |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Tenant Code      : {0,-30}|" -f $TenantCode)         -ForegroundColor White
Write-Host ("  | Scope            : {0,-30}|" -f $Scope)              -ForegroundColor White
Write-Host ("  | Total Users      : {0,-30}|" -f $TotalCount)         -ForegroundColor White
Write-Host ("  |   Enabled        : {0,-30}|" -f $EnabledCount)       -ForegroundColor White
Write-Host ("  |   Disabled       : {0,-30}|" -f $DisabledCount)      -ForegroundColor White
Write-Host ("  | With SMTP        : {0,-30}|" -f $WithSmtpCount)      -ForegroundColor White
Write-Host ("  | With Title       : {0,-30}|" -f $WithTitleCount)     -ForegroundColor White
Write-Host ("  | With Department  : {0,-30}|" -f $WithDeptCount)      -ForegroundColor White
Write-Host ("  | With Manager     : {0,-30}|" -f $WithManagerCount)   -ForegroundColor White
Write-Host ("  | With License     : {0,-30}|" -f $WithLicenseCount)   -ForegroundColor White
Write-Host ("  | With Role        : {0,-30}|" -f $WithRoleCount)      -ForegroundColor White
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Total Run Time   : {0,-30}|" -f $RunTime)            -ForegroundColor Yellow
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""
Read-Host "  Press Enter to exit"
