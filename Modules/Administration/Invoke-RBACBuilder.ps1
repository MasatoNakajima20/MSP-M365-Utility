<#
    RBAC-Builder.ps1

    Interactively provisions Exchange Online RBAC-for-Applications so a Graph app
    (Application permissions) is restricted to a specific set of mailboxes instead
    of the tenant-wide default.

    Flow:
      1. Connect to Microsoft Graph + Exchange Online
      2. Resolve the app (Entra AppId / ObjectId) from a typed App Name
      3. Register the app as an EXO Service Principal
      4. Prompt for a friendly name -> derive Management Scope name
      5. Prompt for mailboxes -> create a mail-enabled security group -> add members
      6. Create the Management Scope (filtered on group membership)
      7. Auto-map the app's consented Graph permissions to matching EXO Application roles
      8. Create the Management Role Assignment
      9. Print the EXO Service Principal Object ID
     10. Test loop: Test-ServicePrincipalAuthorization against typed addresses until "done"

    No #Requires -RunAsAdministrator: every cmdlet here is a Graph/EXO cloud call,
    so local admin rights are not needed (explicitly dropped per sign-off).
#>

$LogPath = "C:\Logging\RBAC-Builder\"
if (-not (Test-Path -Path $LogPath)) {
    New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
}
$LogFile = Join-Path -Path $LogPath -ChildPath ("RBAC-Builder_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

# Writes a timestamped line to both the console and the session log file.
function Write-Log {
    param(
        [Parameter(Mandatory)] [string]$Message,
        [ValidateSet("INFO","WARN","ERROR")] [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $line
    switch ($Level) {
        "WARN"  { Write-Host $line -ForegroundColor Yellow }
        "ERROR" { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line -ForegroundColor Gray }
    }
}

# Prints the cyan banner header used to open the script run.
function Write-Banner {
    param([string]$Title)
    Write-Host ""
    Write-Host ("=" * 70) -ForegroundColor Cyan
    Write-Host $Title -ForegroundColor Cyan
    Write-Host ("=" * 70) -ForegroundColor Cyan
    Write-Host ""
}

# Connects to Microsoft Graph and Exchange Online with the scopes/cmdlets this script needs.
function Connect-RbacBuilderSessions {
    Write-Log "Connecting to Microsoft Graph..."
    Connect-MgGraph -Scopes "Application.Read.All" -NoWelcome
    Write-Log "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false
    Write-Log "Connected to Graph and Exchange Online."
}

# Resolves an app's Entra AppId and Object ID from a typed display name, handling zero/multiple matches.
function Get-EntraAppServicePrincipal {
    param([Parameter(Mandatory)] [string]$AppDisplayName)

    $matches = Get-MgServicePrincipal -Filter "displayName eq '$AppDisplayName'"

    if (-not $matches) {
        Write-Log "No Entra service principal found for '$AppDisplayName'." -Level "ERROR"
        return $null
    }

    if (@($matches).Count -gt 1) {
        Write-Log "Multiple service principals matched '$AppDisplayName'. Select one:" -Level "WARN"
        for ($i = 0; $i -lt $matches.Count; $i++) {
            Write-Host "  [$i] $($matches[$i].DisplayName) - AppId: $($matches[$i].AppId) - ObjectId: $($matches[$i].Id)"
        }
        $selection = Read-Host "Enter the number of the correct app"
        $entraServicePrincipal = $matches[[int]$selection]
    }
    else {
        $entraServicePrincipal = $matches
    }

    return $entraServicePrincipal
}

# Registers the app as an Exchange Online service principal so RBAC-for-Applications can target it.
function New-ExoServicePrincipalForApp {
    param(
        [Parameter(Mandatory)] [string]$EntraAppId,
        [Parameter(Mandatory)] [string]$EntraObjectId,
        [Parameter(Mandatory)] [string]$ExoServicePrincipalDisplayName
    )

    Write-Log "Creating EXO Service Principal '$ExoServicePrincipalDisplayName'..."
    $exoServicePrincipal = New-ServicePrincipal -AppId $EntraAppId -ObjectId $EntraObjectId -DisplayName $ExoServicePrincipalDisplayName
    Write-Log "EXO Service Principal created: $($exoServicePrincipal.ObjectId)"
    return $exoServicePrincipal
}

# Prompts for mailboxes (multiline) and creates/populates a mail-enabled security group for scoping.
function New-MailboxAccessGroup {
    param([Parameter(Mandatory)] [string]$SecurityGroupName)

    Write-Host "Enter mailbox SMTP addresses to allow access to, one per line."
    Write-Host "Press Enter on a blank line when finished."
    $mailboxList = @()
    while ($true) {
        $line = Read-Host "Mailbox"
        if ([string]::IsNullOrWhiteSpace($line)) { break }
        $mailboxList += $line.Trim()
    }

    Write-Log "Creating mail-enabled security group '$SecurityGroupName'..."
    $securityGroup = New-DistributionGroup -Name $SecurityGroupName -Type Security

    foreach ($mailbox in $mailboxList) {
        Write-Log "Adding '$mailbox' to '$SecurityGroupName'..."
        Add-DistributionGroupMember -Identity $SecurityGroupName -Member $mailbox
    }

    return [PSCustomObject]@{
        MailboxList   = $mailboxList
        SecurityGroup = $securityGroup
    }
}

# Creates the Management Scope filtered on membership in the access group.
function New-RbacManagementScope {
    param(
        [Parameter(Mandatory)] [string]$ManagementScopeName,
        [Parameter(Mandatory)] [string]$SecurityGroupName
    )

    $groupDn = (Get-DistributionGroup -Identity $SecurityGroupName).DistinguishedName
    Write-Log "Creating Management Scope '$ManagementScopeName' filtered on group membership..."
    $managementScope = New-ManagementScope -Name $ManagementScopeName -RecipientRestrictionFilter "MemberOfGroup -eq '$groupDn'"
    return $managementScope
}

# Reads the app's consented Microsoft Graph application permissions and maps them to EXO "Application X" role names.
function Get-MappedExoRoles {
    param([Parameter(Mandatory)] [string]$EntraObjectId)

    $roleMap = @{
        "Mail.Read"                   = "Application Mail.Read"
        "Mail.ReadWrite"               = "Application Mail.ReadWrite"
        "Mail.Send"                    = "Application Mail.Send"
        "MailboxSettings.Read"         = "Application MailboxSettings.Read"
        "MailboxSettings.ReadWrite"    = "Application MailboxSettings.ReadWrite"
        "Calendars.Read"               = "Application Calendars.Read"
        "Calendars.ReadWrite"          = "Application Calendars.ReadWrite"
        "Contacts.Read"                = "Application Contacts.Read"
        "Contacts.ReadWrite"           = "Application Contacts.ReadWrite"
        "EWS.AccessAsApp"              = "Application EWS.AccessAsApp"
    }

    $graphServicePrincipal = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
    $assignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $EntraObjectId |
        Where-Object { $_.ResourceId -eq $graphServicePrincipal.Id }

    $consentedGraphPermissions = foreach ($assignment in $assignments) {
        ($graphServicePrincipal.AppRoles | Where-Object { $_.Id -eq $assignment.AppRoleId }).Value
    }

    $mappedExoRoles = foreach ($permission in $consentedGraphPermissions) {
        if ($roleMap.ContainsKey($permission)) {
            $roleMap[$permission]
        }
        else {
            Write-Log "No known EXO role mapping for consented permission '$permission'. Skipping." -Level "WARN"
        }
    }

    return [PSCustomObject]@{
        ConsentedGraphPermissions = $consentedGraphPermissions
        MappedExoRoles            = $mappedExoRoles
    }
}

# Creates the Management Role Assignment tying the EXO Service Principal, mapped role(s), and scope together.
function New-RbacRoleAssignment {
    param(
        [Parameter(Mandatory)] [string]$ExoServicePrincipalObjectId,
        [Parameter(Mandatory)] [string[]]$MappedExoRoles,
        [Parameter(Mandatory)] [string]$ManagementScopeName,
        [Parameter(Mandatory)] [string]$RoleAssignmentNameBase
    )

    foreach ($role in $MappedExoRoles) {
        $roleAssignmentName = "$RoleAssignmentNameBase-$($role -replace '\s','')"
        Write-Log "Creating Management Role Assignment '$roleAssignmentName' for role '$role'..."
        New-ManagementRoleAssignment -Name $roleAssignmentName -App $ExoServicePrincipalObjectId -Role $role -CustomResourceScope $ManagementScopeName | Out-Null
    }
}

# Loops prompting for a mailbox address, tests scope authorization for it, and prints Email/InScope until "done".
function Start-ServicePrincipalTestLoop {
    param([Parameter(Mandatory)] [string]$ExoServicePrincipalObjectId)

    $testResults = @()
    Write-Host ""
    Write-Host "Test loop: type an email to test, or 'done' to exit." -ForegroundColor Cyan

    while ($true) {
        $testEmailInput = Read-Host "Email to test"
        if ($testEmailInput -eq "done") { break }

        $testResult = Test-ServicePrincipalAuthorization -Identity $ExoServicePrincipalObjectId -Resource $testEmailInput
        $inScope = [bool]($testResult | Where-Object { $_.Granted -eq $true })

        $color = if ($inScope) { "Green" } else { "Red" }
        Write-Host "Email: $testEmailInput | InScope: $inScope" -ForegroundColor $color

        $testResults += [PSCustomObject]@{
            Email   = $testEmailInput
            InScope = $inScope
        }
    }

    return $testResults
}

# Writes the test loop results to a color-coded HTML table (True = green, False = red) per standing output convention.
function Export-TestResultsHtml {
    param([Parameter(Mandatory)] [array]$TestResults)

    if (-not $TestResults -or $TestResults.Count -eq 0) { return }

    $rows = foreach ($result in $TestResults) {
        $color = if ($result.InScope) { "#C6EFCE" } else { "#FFC7CE" }
        "<tr><td>$($result.Email)</td><td style='background-color:$color'>$($result.InScope)</td></tr>"
    }

    $html = @"
<html>
<head><style>
table { border-collapse: collapse; font-family: Segoe UI, Arial, sans-serif; }
th, td { border: 1px solid #999; padding: 6px 12px; text-align: left; }
th { background-color: #4472C4; color: white; }
</style></head>
<body>
<h3>RBAC Builder - Test Loop Results</h3>
<table>
<tr><th>Email</th><th>InScope</th></tr>
$($rows -join "`n")
</table>
</body>
</html>
"@

    $htmlPath = Join-Path -Path $LogPath -ChildPath ("TestResults_{0}.html" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
    $html | Out-File -FilePath $htmlPath -Encoding UTF8
    Write-Log "Test results exported to $htmlPath"
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

Write-Banner "RBAC Builder"

Connect-RbacBuilderSessions

$AppDisplayName = Read-Host "App Name"
$EntraServicePrincipal = Get-EntraAppServicePrincipal -AppDisplayName $AppDisplayName
if (-not $EntraServicePrincipal) {
    Write-Log "Exiting: could not resolve app '$AppDisplayName'." -Level "ERROR"
    exit 0
}
$EntraAppId   = $EntraServicePrincipal.AppId
$EntraObjectId = $EntraServicePrincipal.Id

$ExoServicePrincipalDisplayName = $AppDisplayName + "_Service_Principal"
$ExoServicePrincipal = New-ExoServicePrincipalForApp -EntraAppId $EntraAppId -EntraObjectId $EntraObjectId -ExoServicePrincipalDisplayName $ExoServicePrincipalDisplayName
$ExoServicePrincipalObjectId = $ExoServicePrincipal.ObjectId

$FriendlyName = Read-Host "Name"
$ManagementScopeName = $FriendlyName + "_Management_Scope"

$SecurityGroupName = $AppDisplayName + "_RBAC_Access"
$groupResult = New-MailboxAccessGroup -SecurityGroupName $SecurityGroupName
$MailboxList = $groupResult.MailboxList
$SecurityGroup = $groupResult.SecurityGroup

New-RbacManagementScope -ManagementScopeName $ManagementScopeName -SecurityGroupName $SecurityGroupName | Out-Null

$roleResult = Get-MappedExoRoles -EntraObjectId $EntraObjectId
$ConsentedGraphPermissions = $roleResult.ConsentedGraphPermissions
$MappedExoRoles = $roleResult.MappedExoRoles

if (-not $MappedExoRoles -or $MappedExoRoles.Count -eq 0) {
    Write-Log "No consented Graph permissions mapped to a known EXO role. Skipping role assignment." -Level "WARN"
}
else {
    $RoleAssignmentName = $FriendlyName + "_RoleAssignment"
    New-RbacRoleAssignment -ExoServicePrincipalObjectId $ExoServicePrincipalObjectId -MappedExoRoles $MappedExoRoles -ManagementScopeName $ManagementScopeName -RoleAssignmentNameBase $RoleAssignmentName
}

Write-Host ""
Write-Host "EXO Service Principal Object ID: $ExoServicePrincipalObjectId" -ForegroundColor Cyan
Write-Host ""

$TestResults = Start-ServicePrincipalTestLoop -ExoServicePrincipalObjectId $ExoServicePrincipalObjectId
Export-TestResultsHtml -TestResults $TestResults

Write-Log "RBAC Builder run complete."
exit 0
