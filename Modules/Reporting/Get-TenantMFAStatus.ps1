#Requires -Modules Microsoft.Graph.Users, Microsoft.Graph.Reports, Microsoft.Graph.Identity.SignIns

<#
.SYNOPSIS
    Reports MFA registration and authentication methods for all licensed,
    non-guest users in a Microsoft 365 tenant.

.DESCRIPTION
    Connects to Microsoft Graph and enumerates all licensed Member users
    (guests excluded). For each user, the script reports:
        - MFA registration status (registered / capable)
        - Default MFA method (as configured in the tenant)
        - Priority method (Microsoft Authenticator > SMS/Voice > Software OTP)
        - Booleans for each method family:
            * HasMSAuthenticator
            * HasSMSorVoice
            * HasSoftwareOTP
        - All registered methods (comma-separated)

    The script prefers the Authentication Methods user-registration report
    for speed and falls back to per-user method enumeration if the report
    is unavailable.

    Results are exported to C:\MSP-M365-Utility\.

.NOTES
    Required Modules:
        - Microsoft.Graph.Users
        - Microsoft.Graph.Reports
        - Microsoft.Graph.Identity.SignIns
    Required Permissions (Graph):
        - User.Read.All
        - UserAuthenticationMethod.Read.All
        - AuditLog.Read.All     (for the registration report)
#>

# ----------------------------------------------------------
#  CONFIGURATION
# ----------------------------------------------------------
$ErrorActionPreference = 'Stop'
$Host.UI.RawUI.WindowTitle = "M365 MFA Status Report"

# ----------------------------------------------------------
#  BANNER
# ----------------------------------------------------------
Clear-Host
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |          M365 MFA Status Report Script           |" -ForegroundColor Cyan
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
#  OUTPUT PATH (C:\MSP-M365-Utility\)
# ----------------------------------------------------------
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
$OutputFile = "MFAStatus_${TenantCode}_${Timestamp}.csv"
$OutputPath = Join-Path $OutputRoot $OutputFile

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
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Reports',
    'Microsoft.Graph.Identity.SignIns'
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
#  CONNECT TO GRAPH
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [2/4] Connecting to Microsoft Graph..." -ForegroundColor Cyan

try {
    Connect-MgGraph -Scopes "User.Read.All","UserAuthenticationMethod.Read.All","AuditLog.Read.All" -NoWelcome -ErrorAction Stop
    Write-Host "  [OK] Microsoft Graph connected." -ForegroundColor Green
}
catch {
    Write-Host "  [ERROR] Failed to connect: $_" -ForegroundColor Red
    Read-Host "`n  Press Enter to exit"
    exit 1
}

# ----------------------------------------------------------
#  RETRIEVE LICENSED MEMBER USERS (GUESTS EXCLUDED)
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [3/4] Retrieving licensed users (guests excluded)..." -ForegroundColor Cyan

try {
    $AllMembers = Get-MgUser -All `
        -Filter "userType eq 'Member'" `
        -Property "Id,DisplayName,UserPrincipalName,Mail,AssignedLicenses,AccountEnabled,UserType" `
        -ErrorAction Stop

    $Users      = $AllMembers | Where-Object { $_.AssignedLicenses -and $_.AssignedLicenses.Count -gt 0 }
    $TotalCount = @($Users).Count
    Write-Host "  [OK] Found $TotalCount licensed user(s)." -ForegroundColor Green
}
catch {
    Write-Host "  [ERROR] Could not retrieve users: $_" -ForegroundColor Red
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Read-Host "`n  Press Enter to exit"
    exit 1
}

if ($TotalCount -eq 0) {
    Write-Host "  [!] No licensed users found. Nothing to report." -ForegroundColor Yellow
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Read-Host "`n  Press Enter to exit"
    exit 0
}

# ----------------------------------------------------------
#  PRE-FETCH AUTHENTICATION METHOD REGISTRATION REPORT
# ----------------------------------------------------------
Write-Host "  Fetching authentication method registration report..." -ForegroundColor DarkCyan

$RegLookup    = @{}
$ReportLoaded = $false
try {
    Get-MgReportAuthenticationMethodUserRegistrationDetail -All -ErrorAction Stop |
        ForEach-Object {
            if ($_.UserPrincipalName) {
                $RegLookup[$_.UserPrincipalName.ToLower()] = $_
            }
        }
    $ReportLoaded = $true
    Write-Host "  [OK] Registration report cached ($($RegLookup.Count) entries)." -ForegroundColor Green
}
catch {
    Write-Host "  [WARN] Registration report unavailable: $_" -ForegroundColor Yellow
    Write-Host "  [WARN] Falling back to per-user method queries (slower)." -ForegroundColor Yellow
}

# ----------------------------------------------------------
#  HELPER: CLASSIFY METHOD STRINGS
# ----------------------------------------------------------
function Get-MethodFlags {
    param ([string[]]$Methods)

    $flags = [ordered]@{
        HasMSAuthenticator = $false
        HasSMSorVoice      = $false
        HasSoftwareOTP     = $false
        Other              = [System.Collections.Generic.List[string]]::new()
    }
    if (-not $Methods) { return $flags }

    foreach ($m in $Methods) {
        if (-not $m) { continue }
        switch -Regex ($m) {
            'microsoftAuthenticator|passwordlessMicrosoftAuthenticator|MicrosoftAuthenticator' {
                $flags.HasMSAuthenticator = $true ; continue
            }
            'mobilePhone|alternateMobilePhone|officePhone|sms|voice|phoneAuthentication' {
                $flags.HasSMSorVoice = $true ; continue
            }
            'softwareOath|softwareOneTimePasscode|oath' {
                $flags.HasSoftwareOTP = $true ; continue
            }
            default {
                [void]$flags.Other.Add($m)
            }
        }
    }
    return $flags
}

# Strip Graph @odata.type into a friendly method name
function ConvertTo-MethodName {
    param ([string]$ODataType)
    if (-not $ODataType) { return $null }
    $leaf = ($ODataType -split '\.')[-1]
    return ($leaf -replace 'AuthenticationMethod$','')
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

    Write-Progress `
        -Activity        "Processing Users -- Tenant: $TenantCode" `
        -Status          "$ProgressBar  $Counter / $TotalCount  ($PercentComplete%)  |  $($User.DisplayName)" `
        -PercentComplete  $PercentComplete `
        -CurrentOperation "Current: $($User.UserPrincipalName)"

    $MfaRegistered = $null
    $MfaCapable    = $null
    $DefaultMethod = $null
    $Methods       = @()

    $Reg = $null
    if ($User.UserPrincipalName) {
        $Reg = $RegLookup[$User.UserPrincipalName.ToLower()]
    }

    if ($Reg) {
        $MfaRegistered = $Reg.IsMfaRegistered
        $MfaCapable    = $Reg.IsMfaCapable
        $DefaultMethod = $Reg.DefaultMfaMethod
        $Methods       = @($Reg.MethodsRegistered)
    }
    else {
        # Fallback: enumerate methods per user
        try {
            $UserMethods = Get-MgUserAuthenticationMethod -UserId $User.Id -ErrorAction Stop
            $Methods = @(
                $UserMethods | ForEach-Object { ConvertTo-MethodName $_.AdditionalProperties['@odata.type'] } |
                    Where-Object { $_ -and $_ -notin @('password') }
            )
            $MfaRegistered = ($Methods.Count -gt 0)
            $MfaCapable    = $MfaRegistered
        }
        catch {
            $MfaRegistered = $null
            $MfaCapable    = $null
        }
    }

    $flags = Get-MethodFlags -Methods $Methods

    # Priority: MS Authenticator > SMS/Voice > Software OTP > Other > None
    $PriorityMethod =
        if     ($flags.HasMSAuthenticator)  { 'Microsoft Authenticator' }
        elseif ($flags.HasSMSorVoice)       { 'SMS / Voice'             }
        elseif ($flags.HasSoftwareOTP)      { 'Software OTP'            }
        elseif ($flags.Other.Count -gt 0)   { ($flags.Other -join ', ') }
        else                                { 'None'                    }

    $Results.Add([PSCustomObject]@{
        DisplayName        = $User.DisplayName
        UserPrincipalName  = $User.UserPrincipalName
        PrimarySmtpAddress = if ($User.Mail) { $User.Mail } else { $User.UserPrincipalName }
        AccountEnabled     = $User.AccountEnabled
        MFARegistered      = $MfaRegistered
        MFACapable         = $MfaCapable
        DefaultMethod      = $DefaultMethod
        PriorityMethod     = $PriorityMethod
        HasMSAuthenticator = $flags.HasMSAuthenticator
        HasSMSorVoice      = $flags.HasSMSorVoice
        HasSoftwareOTP     = $flags.HasSoftwareOTP
        AllMethods         = ($Methods -join ', ')
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

$MfaYes        = ($Results | Where-Object { $_.MFARegistered -eq $true  }).Count
$MfaNo         = ($Results | Where-Object { $_.MFARegistered -eq $false }).Count
$MfaUnknown    = ($Results | Where-Object { $null -eq $_.MFARegistered  }).Count
$MsAuthCount   = ($Results | Where-Object { $_.HasMSAuthenticator -eq $true }).Count
$SmsVoiceCount = ($Results | Where-Object { $_.HasSMSorVoice      -eq $true }).Count
$SwOtpCount    = ($Results | Where-Object { $_.HasSoftwareOTP     -eq $true }).Count
$NoMethodCount = ($Results | Where-Object { $_.PriorityMethod -eq 'None' }).Count

$RunTime = "{0:D2}h {1:D2}m {2:D2}s {3:D3}ms" -f `
    $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds, $Elapsed.Milliseconds

Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |                  RUN SUMMARY                    |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Tenant Code            : {0,-24}|" -f $TenantCode)    -ForegroundColor White
Write-Host ("  | Source                 : {0,-24}|" -f $(if ($ReportLoaded) {'Registration Report'} else {'Per-User Query'})) -ForegroundColor White
Write-Host ("  | Total Licensed Users   : {0,-24}|" -f $TotalCount)    -ForegroundColor White
Write-Host ("  |   MFA Registered       : {0,-24}|" -f $MfaYes)        -ForegroundColor White
Write-Host ("  |   Not Registered       : {0,-24}|" -f $MfaNo)         -ForegroundColor White
Write-Host ("  |   Unknown              : {0,-24}|" -f $MfaUnknown)    -ForegroundColor White
Write-Host "  |                                                  |" -ForegroundColor Cyan
Write-Host ("  |   MS Authenticator     : {0,-24}|" -f $MsAuthCount)   -ForegroundColor White
Write-Host ("  |   SMS / Voice          : {0,-24}|" -f $SmsVoiceCount) -ForegroundColor White
Write-Host ("  |   Software OTP         : {0,-24}|" -f $SwOtpCount)    -ForegroundColor White
Write-Host ("  |   No Method            : {0,-24}|" -f $NoMethodCount) -ForegroundColor White
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("  | Total Run Time         : {0,-24}|" -f $RunTime)       -ForegroundColor Yellow
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""
Read-Host "  Press Enter to exit"
