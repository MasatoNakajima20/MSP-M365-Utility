#Requires -Version 5.1

<#
.SYNOPSIS
    MSP M365 Utility - GUI launcher for the modules in this repo.

.DESCRIPTION
    Single-file WinForms launcher. Each module is downloaded from GitHub on
    demand into %TEMP%\MSPM365Utility and started in its own PowerShell
    window so the module's Read-Host prompts work normally.

.EXAMPLE
    # Run on any Windows machine - no local clone needed:
    iex (irm "https://raw.githubusercontent.com/MasatoNakajima20/MSP-M365-Utility/main/Launch-MSPM365Utility.ps1")

.NOTES
    Repo : https://github.com/MasatoNakajima20/MSP-M365-Utility
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:RepoOwner  = 'MasatoNakajima20'
$script:RepoName   = 'MSP-M365-Utility'
$script:Branch     = 'main'
$script:Version    = '0.9.0-beta'
$script:BaseRawUrl = "https://raw.githubusercontent.com/$script:RepoOwner/$script:RepoName/$script:Branch"
$script:WorkDir    = Join-Path $env:TEMP 'MSPM365Utility'   # module cache (internal)
$script:ResultsDir = 'C:\MSP-M365-Utility'                  # where reporting modules drop CSVs

# Module catalog. When a new module is added to /Modules, add it here too.
$script:Modules = @(
    [PSCustomObject]@{
        File        = 'Modules/Reporting/Get-TenantMailboxes.ps1'
        Title       = 'Tenant Mailbox Inventory'
        Category    = 'Reporting'
        Description = 'Export every mailbox (user, shared, resource) with enabled and licensed status to CSV.'
    }
    [PSCustomObject]@{
        File        = 'Modules/Reporting/Get-TenantUserDetails.ps1'
        Title       = 'Tenant User Details'
        Category    = 'Reporting'
        Description = 'Export user details - name, email, title, department, manager - to CSV.'
    }
    [PSCustomObject]@{
        File        = 'Modules/Reporting/Get-TenantGroupMembership.ps1'
        Title       = 'Tenant Group Membership'
        Category    = 'Reporting'
        Description = 'Export all groups (DL, mail-enabled security, M365, security) and members to CSV.'
    }
    [PSCustomObject]@{
        File        = 'Modules/Reporting/Get-TenantMFAStatus.ps1'
        Title       = 'Tenant MFA Status'
        Category    = 'Reporting'
        Description = 'Licensed users (guests excluded) with MFA status, default method, and Authenticator/SMS/OTP flags.'
    }
    [PSCustomObject]@{
        File        = 'Modules/Administration/Add-DistroMember.ps1'
        Title       = 'Add Distribution List Members'
        Category    = 'Administration'
        Description = 'Bulk add members to a distribution list from a pasted email list.'
    }
    [PSCustomObject]@{
        File        = 'Modules/Administration/Remove-DistroMember.ps1'
        Title       = 'Remove Distribution List Members'
        Category    = 'Administration'
        Description = 'Bulk remove members from a distribution list from a pasted email list.'
    }
    [PSCustomObject]@{
        File        = 'Modules/Administration/Add-BulkContact.ps1'
        Title       = 'Bulk Contact Creation'
        Category    = 'Administration'
        Description = 'Bulk create external Mail Contacts from a paste-list. Prompts per duplicate.'
    }
    [PSCustomObject]@{
        File        = 'Modules/Administration/Add-BulkGuestUser.ps1'
        Title       = 'Bulk Guest User Invite'
        Category    = 'Administration'
        Description = 'Bulk send B2B guest invitations. Auto-detects tenant redirect URL. Prompts per duplicate.'
    }
    [PSCustomObject]@{
        File        = 'Modules/Administration/Add-CalendarPermission.ps1'
        Title       = 'Add Calendar Permission'
        Category    = 'Administration'
        Description = 'Grant Reviewer / Author / Editor access to a mailbox calendar. Role pickable per user.'
    }
    [PSCustomObject]@{
        File        = 'Modules/Administration/Remove-CalendarPermission.ps1'
        Title       = 'Remove Calendar Permission'
        Category    = 'Administration'
        Description = 'Remove user(s) calendar access on a target mailbox. One confirmation, then straight removal.'
    }
    [PSCustomObject]@{
        File        = 'Modules/Administration/Add-MailboxAccess.ps1'
        Title       = 'Add Mailbox Access (FullAccess + SendAs)'
        Category    = 'Administration'
        Description = 'Bulk grant FullAccess + SendAs. Creates target as a shared mailbox if it does not exist.'
    }
    [PSCustomObject]@{
        File        = 'Modules/Administration/Remove-MailboxAccess.ps1'
        Title       = 'Remove Mailbox Access (FullAccess + SendAs)'
        Category    = 'Administration'
        Description = 'Bulk remove FullAccess + SendAs from a target mailbox.'
    }
    [PSCustomObject]@{
        File        = 'Modules/Administration/Remove-UserAccess.ps1'
        Title       = 'User Access: Discover and Remove'
        Category    = 'Administration'
        Description = 'Audit a user''s mailbox access and group memberships, then selectively remove (paste-list or "all").'
    }
    [PSCustomObject]@{
        File        = 'Modules/Administration/Request-OneDriveProvision.ps1'
        Title       = 'OneDrive Pre-Provisioning'
        Category    = 'Administration'
        Description = 'Pre-provision OneDrive for a pasted list of users, one by one. SharePoint Online.'
    }
    [PSCustomObject]@{
        File        = 'Modules/Utility/Install-ExchangeOnlineModule.ps1'
        Title       = 'Install Exchange Online Module'
        Category    = 'Utility'
        Description = 'Install / update the ExchangeOnlineManagement PowerShell module (CurrentUser scope).'
    }
    [PSCustomObject]@{
        File        = 'Modules/Utility/Install-MicrosoftGraphModule.ps1'
        Title       = 'Install Microsoft Graph Modules'
        Category    = 'Utility'
        Description = 'Install / update the 4 Microsoft.Graph submodules used by these scripts (Users, Groups, Reports, Identity.SignIns).'
    }
    [PSCustomObject]@{
        File        = 'Modules/Utility/Install-PowerShell7.ps1'
        Title       = 'Install PowerShell 7'
        Category    = 'Utility'
        Description = 'Install PowerShell 7 via winget (Microsoft.PowerShell package).'
    }
    [PSCustomObject]@{
        File        = 'Modules/Utility/Install-All.ps1'
        Title       = 'Install All Prerequisites'
        Category    = 'Utility'
        Description = 'Run all three installers in sequence. Idempotent - already-installed items are skipped.'
    }
)

# Brand palette
$BrandBlue       = [System.Drawing.ColorTranslator]::FromHtml('#008BC7')
$BrandBlueDark   = [System.Drawing.ColorTranslator]::FromHtml('#00577E')
$BrandBlueLight  = [System.Drawing.ColorTranslator]::FromHtml('#E6F4FA')
$BrandTextDark   = [System.Drawing.ColorTranslator]::FromHtml('#262626')
$BrandTextMuted  = [System.Drawing.ColorTranslator]::FromHtml('#999999')

function Ensure-WorkDir {
    if (-not (Test-Path $script:WorkDir)) {
        New-Item -ItemType Directory -Path $script:WorkDir -Force | Out-Null
    }
}

function Invoke-Module {
    param([PSCustomObject]$Module, [System.Windows.Forms.Label]$StatusLabel)

    Ensure-WorkDir
    $url       = "$script:BaseRawUrl/$($Module.File)"
    $leaf      = Split-Path $Module.File -Leaf
    $localPath = Join-Path $script:WorkDir $leaf

    $StatusLabel.Text = "Downloading $leaf ..."
    $StatusLabel.Refresh()

    try {
        Invoke-WebRequest -Uri $url -OutFile $localPath -UseBasicParsing -ErrorAction Stop
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Could not download module from:`n$url`n`n$($_.Exception.Message)",
            'Download Error',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        $StatusLabel.Text = 'Ready.'
        return
    }

    $psExe = if (Get-Command pwsh -ErrorAction SilentlyContinue) { 'pwsh' } else { 'powershell' }
    try {
        # No -NoExit: the window closes when the module finishes. Each module ends
        # with a "Press Enter to exit" pause (including on error/abort) so results
        # stay visible until the operator dismisses the window.
        Start-Process -FilePath $psExe -ArgumentList @(
            '-ExecutionPolicy', 'Bypass',
            '-File', $localPath
        ) | Out-Null
        $StatusLabel.Text = "Launched: $($Module.Title)  (results in $script:ResultsDir)"
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Could not launch $psExe.`n`n$($_.Exception.Message)",
            'Launch Error',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        $StatusLabel.Text = 'Ready.'
    }
}

function Open-WorkDir {
    Ensure-WorkDir
    Start-Process explorer.exe $script:WorkDir | Out-Null
}

function Get-RequiredPrereqStatus {
    # Returns an ordered list of @{Name, Ok, Detail} for the landing-page indicator.
    $result = New-Object System.Collections.Generic.List[PSCustomObject]

    # Exchange Online
    $exo = Get-Module -ListAvailable -Name 'ExchangeOnlineManagement' -ErrorAction SilentlyContinue
    $result.Add([PSCustomObject]@{
        Name   = 'Exchange Online'
        Ok     = [bool]$exo
        Detail = if ($exo) { 'v' + (($exo | Sort-Object Version -Descending | Select-Object -First 1).Version) } else { 'Missing' }
    })

    # MS Graph - require all 4 submodules we actually use
    $graphMods = @('Microsoft.Graph.Users','Microsoft.Graph.Groups','Microsoft.Graph.Reports','Microsoft.Graph.Identity.SignIns')
    $missingGraph = @($graphMods | Where-Object { -not (Get-Module -ListAvailable -Name $_ -ErrorAction SilentlyContinue) })
    $result.Add([PSCustomObject]@{
        Name   = 'MS Graph'
        Ok     = ($missingGraph.Count -eq 0)
        Detail = if ($missingGraph.Count -eq 0) { 'All 4 submodules present' } else { "Missing: $($missingGraph.Count)/$($graphMods.Count)" }
    })

    # PowerShell 7
    $pwsh = Get-Command pwsh -ErrorAction SilentlyContinue
    $result.Add([PSCustomObject]@{
        Name   = 'PowerShell 7'
        Ok     = [bool]$pwsh
        Detail = if ($pwsh) { 'pwsh found' } else { 'Missing' }
    })

    return $result
}

function Open-ResultsDir {
    if (-not (Test-Path $script:ResultsDir)) {
        try {
            New-Item -ItemType Directory -Path $script:ResultsDir -Force -ErrorAction Stop | Out-Null
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Could not create '$script:ResultsDir'.`n`n$($_.Exception.Message)`n`nRun a reporting module first, or create the folder manually.",
                'Results Folder',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }
    }
    Start-Process explorer.exe $script:ResultsDir | Out-Null
}

function Show-AboutDialog {
    $about               = New-Object System.Windows.Forms.Form
    $about.Text          = 'About MSP M365 Utility'
    $about.Size          = New-Object System.Drawing.Size(460, 320)
    $about.StartPosition = 'CenterParent'
    $about.FormBorderStyle = 'FixedDialog'
    $about.MinimizeBox   = $false
    $about.MaximizeBox   = $false
    $about.BackColor     = [System.Drawing.Color]::White
    $about.Font          = New-Object System.Drawing.Font('Segoe UI', 9)

    $hdr           = New-Object System.Windows.Forms.Panel
    $hdr.Size      = New-Object System.Drawing.Size(460, 60)
    $hdr.Location  = New-Object System.Drawing.Point(0, 0)
    $hdr.BackColor = $BrandBlue
    $about.Controls.Add($hdr)

    $hdrTitle           = New-Object System.Windows.Forms.Label
    $hdrTitle.Text      = 'MSP M365 Utility'
    $hdrTitle.Font      = New-Object System.Drawing.Font('Segoe UI', 14, [System.Drawing.FontStyle]::Bold)
    $hdrTitle.ForeColor = [System.Drawing.Color]::White
    $hdrTitle.Location  = New-Object System.Drawing.Point(18, 8)
    $hdrTitle.Size      = New-Object System.Drawing.Size(420, 26)
    $hdrTitle.BackColor = [System.Drawing.Color]::Transparent
    $hdr.Controls.Add($hdrTitle)

    $hdrVer           = New-Object System.Windows.Forms.Label
    $hdrVer.Text      = "Version $script:Version"
    $hdrVer.Font      = New-Object System.Drawing.Font('Segoe UI', 9)
    $hdrVer.ForeColor = [System.Drawing.Color]::White
    $hdrVer.Location  = New-Object System.Drawing.Point(20, 36)
    $hdrVer.Size      = New-Object System.Drawing.Size(420, 18)
    $hdrVer.BackColor = [System.Drawing.Color]::Transparent
    $hdr.Controls.Add($hdrVer)

    $descLbl          = New-Object System.Windows.Forms.Label
    $descLbl.Text     = "GUI launcher for the MSP M365 Utility module set. Fetches each module from GitHub on demand and runs it in its own PowerShell window."
    $descLbl.Font     = New-Object System.Drawing.Font('Segoe UI', 9)
    $descLbl.ForeColor= $BrandTextDark
    $descLbl.Location = New-Object System.Drawing.Point(20, 76)
    $descLbl.Size     = New-Object System.Drawing.Size(415, 44)
    $about.Controls.Add($descLbl)

    $repoCaption          = New-Object System.Windows.Forms.Label
    $repoCaption.Text     = 'Repository:'
    $repoCaption.Font     = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $repoCaption.ForeColor= $BrandTextDark
    $repoCaption.Location = New-Object System.Drawing.Point(20, 130)
    $repoCaption.Size     = New-Object System.Drawing.Size(100, 18)
    $about.Controls.Add($repoCaption)

    $repoUrl  = "https://github.com/$script:RepoOwner/$script:RepoName"
    $repoLink           = New-Object System.Windows.Forms.LinkLabel
    $repoLink.Text      = $repoUrl
    $repoLink.Font      = New-Object System.Drawing.Font('Segoe UI', 9)
    $repoLink.LinkColor = $BrandBlue
    $repoLink.ActiveLinkColor = $BrandBlueDark
    $repoLink.Location  = New-Object System.Drawing.Point(20, 150)
    $repoLink.Size      = New-Object System.Drawing.Size(415, 18)
    $repoLink.Add_LinkClicked({ Start-Process $repoUrl | Out-Null }.GetNewClosure())
    $about.Controls.Add($repoLink)

    $relCaption          = New-Object System.Windows.Forms.Label
    $relCaption.Text     = 'Releases:'
    $relCaption.Font     = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $relCaption.ForeColor= $BrandTextDark
    $relCaption.Location = New-Object System.Drawing.Point(20, 175)
    $relCaption.Size     = New-Object System.Drawing.Size(100, 18)
    $about.Controls.Add($relCaption)

    $relUrl   = "https://github.com/$script:RepoOwner/$script:RepoName/releases"
    $relLink           = New-Object System.Windows.Forms.LinkLabel
    $relLink.Text      = $relUrl
    $relLink.Font      = New-Object System.Drawing.Font('Segoe UI', 9)
    $relLink.LinkColor = $BrandBlue
    $relLink.ActiveLinkColor = $BrandBlueDark
    $relLink.Location  = New-Object System.Drawing.Point(20, 195)
    $relLink.Size      = New-Object System.Drawing.Size(415, 18)
    $relLink.Add_LinkClicked({ Start-Process $relUrl | Out-Null }.GetNewClosure())
    $about.Controls.Add($relLink)

    $okBtn           = New-Object System.Windows.Forms.Button
    $okBtn.Text      = 'OK'
    $okBtn.Size      = New-Object System.Drawing.Size(90, 30)
    $okBtn.Location  = New-Object System.Drawing.Point(335, 235)
    $okBtn.BackColor = $BrandBlue
    $okBtn.ForeColor = [System.Drawing.Color]::White
    $okBtn.FlatStyle = 'Flat'
    $okBtn.FlatAppearance.BorderSize = 0
    $okBtn.Font      = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $okBtn.Add_Click({ $about.Close() })
    $about.AcceptButton = $okBtn
    $about.Controls.Add($okBtn)

    [void]$about.ShowDialog()
    $about.Dispose()
}

# ----- form -----
$form               = New-Object System.Windows.Forms.Form
$form.Text          = "MSP M365 Utility  -  $script:Version"
$form.Size          = New-Object System.Drawing.Size(1100, 620)
$form.AutoScaleMode = 'Dpi'
$form.StartPosition = 'CenterScreen'
$form.BackColor     = [System.Drawing.Color]::White
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox   = $false
$form.Font          = New-Object System.Drawing.Font('Segoe UI', 9)

# ----- header -----
$header           = New-Object System.Windows.Forms.Panel
$header.Size      = New-Object System.Drawing.Size(1100, 72)
$header.Location  = New-Object System.Drawing.Point(0, 0)
$header.BackColor = $BrandBlue
$form.Controls.Add($header)

$titleLbl          = New-Object System.Windows.Forms.Label
$titleLbl.Text     = 'MSP M365 Utility'
$titleLbl.Font     = New-Object System.Drawing.Font('Segoe UI', 18, [System.Drawing.FontStyle]::Bold)
$titleLbl.ForeColor= [System.Drawing.Color]::White
$titleLbl.Location = New-Object System.Drawing.Point(20, 10)
$titleLbl.Size     = New-Object System.Drawing.Size(600, 30)
$titleLbl.BackColor= [System.Drawing.Color]::Transparent
$header.Controls.Add($titleLbl)

$subLbl          = New-Object System.Windows.Forms.Label
$subLbl.Text     = 'Choose a category to see its modules. Each module opens in its own PowerShell window.'
$subLbl.Font     = New-Object System.Drawing.Font('Segoe UI', 9)
$subLbl.ForeColor= [System.Drawing.Color]::White
$subLbl.Location = New-Object System.Drawing.Point(22, 42)
$subLbl.Size     = New-Object System.Drawing.Size(700, 20)
$subLbl.BackColor= [System.Drawing.Color]::Transparent
$header.Controls.Add($subLbl)

# Forward-declare so card / tile click handlers can reach it
$statusLbl = New-Object System.Windows.Forms.Label

# ----- category-selection panel (shown first) -----
$categoryPanel          = New-Object System.Windows.Forms.Panel
$categoryPanel.Location = New-Object System.Drawing.Point(15, 85)
$categoryPanel.Size     = New-Object System.Drawing.Size(1065, 430)
$categoryPanel.BackColor= [System.Drawing.Color]::White
$form.Controls.Add($categoryPanel)

# ----- modules panel (hidden until a category is picked) -----
$modulesPanel           = New-Object System.Windows.Forms.Panel
$modulesPanel.Location  = New-Object System.Drawing.Point(15, 85)
$modulesPanel.Size      = New-Object System.Drawing.Size(1065, 430)
$modulesPanel.BackColor = [System.Drawing.Color]::White
$modulesPanel.Visible   = $false
$form.Controls.Add($modulesPanel)

$backBtn           = New-Object System.Windows.Forms.Button
$backBtn.Text      = "< Back to Categories"
$backBtn.Size      = New-Object System.Drawing.Size(180, 28)
$backBtn.Location  = New-Object System.Drawing.Point(0, 4)
$backBtn.FlatStyle = 'Flat'
$backBtn.BackColor = [System.Drawing.Color]::White
$backBtn.ForeColor = $BrandBlueDark
$backBtn.Font      = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$backBtn.Cursor    = [System.Windows.Forms.Cursors]::Hand
$modulesPanel.Controls.Add($backBtn)

$breadcrumbLbl          = New-Object System.Windows.Forms.Label
$breadcrumbLbl.Text     = ''
$breadcrumbLbl.Font     = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
$breadcrumbLbl.ForeColor= $BrandTextDark
$breadcrumbLbl.Location = New-Object System.Drawing.Point(190, 8)
$breadcrumbLbl.Size     = New-Object System.Drawing.Size(500, 22)
$modulesPanel.Controls.Add($breadcrumbLbl)

$list                 = New-Object System.Windows.Forms.FlowLayoutPanel
$list.Location        = New-Object System.Drawing.Point(0, 40)
$list.Size            = New-Object System.Drawing.Size(1065, 388)
$list.FlowDirection   = 'TopDown'
$list.WrapContents    = $false
$list.AutoScroll      = $true
$list.BackColor       = [System.Drawing.Color]::White
$modulesPanel.Controls.Add($list)

# ----- factory: build one module card panel -----
function New-ModuleCard {
    param (
        [PSCustomObject]$Module,
        [System.Windows.Forms.Label]$StatusLabel
    )

    $card             = New-Object System.Windows.Forms.Panel
    $card.Size        = New-Object System.Drawing.Size(1040, 76)
    $card.Margin      = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
    $card.BackColor   = $BrandBlueLight
    $card.BorderStyle = 'FixedSingle'

    $catLbl           = New-Object System.Windows.Forms.Label
    $catLbl.Text      = $Module.Category.ToUpper()
    $catLbl.Font      = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)
    $catLbl.ForeColor = $BrandBlueDark
    $catLbl.Location  = New-Object System.Drawing.Point(12, 8)
    $catLbl.Size      = New-Object System.Drawing.Size(120, 14)
    $catLbl.BackColor = [System.Drawing.Color]::Transparent
    $card.Controls.Add($catLbl)

    $modTitleLbl          = New-Object System.Windows.Forms.Label
    $modTitleLbl.Text     = $Module.Title
    $modTitleLbl.Font     = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
    $modTitleLbl.ForeColor= $BrandTextDark
    $modTitleLbl.Location = New-Object System.Drawing.Point(12, 22)
    $modTitleLbl.Size     = New-Object System.Drawing.Size(840, 22)
    $modTitleLbl.BackColor= [System.Drawing.Color]::Transparent
    $card.Controls.Add($modTitleLbl)

    $descLbl          = New-Object System.Windows.Forms.Label
    $descLbl.Text     = $Module.Description
    $descLbl.Font     = New-Object System.Drawing.Font('Segoe UI', 9)
    $descLbl.ForeColor= $BrandTextDark
    $descLbl.Location = New-Object System.Drawing.Point(12, 46)
    $descLbl.Size     = New-Object System.Drawing.Size(840, 24)
    $descLbl.BackColor= [System.Drawing.Color]::Transparent
    $card.Controls.Add($descLbl)

    $runBtn           = New-Object System.Windows.Forms.Button
    $runBtn.Text      = 'Run'
    $runBtn.Size      = New-Object System.Drawing.Size(140, 40)
    $runBtn.Location  = New-Object System.Drawing.Point(880, 18)
    $runBtn.BackColor = $BrandBlue
    $runBtn.ForeColor = [System.Drawing.Color]::White
    $runBtn.FlatStyle = 'Flat'
    $runBtn.FlatAppearance.BorderSize = 0
    $runBtn.Font      = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
    $runBtn.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $localMod         = $Module
    $localStatus      = $StatusLabel
    $runBtn.Add_Click({ Invoke-Module -Module $localMod -StatusLabel $localStatus }.GetNewClosure())
    $card.Controls.Add($runBtn)

    return $card
}

# ----- view-switching helpers -----
function Show-Categories {
    $modulesPanel.Visible  = $false
    $categoryPanel.Visible = $true
    $statusLbl.Text        = 'Ready.'
}

function Show-Category {
    param ([string]$Category)

    $breadcrumbLbl.Text = $Category.ToUpper()
    $list.Controls.Clear()
    $matches = @($script:Modules | Where-Object { $_.Category -eq $Category })
    foreach ($mod in $matches) {
        $card = New-ModuleCard -Module $mod -StatusLabel $statusLbl
        $list.Controls.Add($card)
    }

    $categoryPanel.Visible = $false
    $modulesPanel.Visible  = $true
    $statusLbl.Text        = "$Category - $($matches.Count) module(s) available."
}

$backBtn.Add_Click({ Show-Categories })

# ----- build category-selection tiles -----
function New-CategoryTile {
    param (
        [string]$Category,
        [string]$Tagline,
        [int]$ModuleCount,
        [int]$X,
        [int]$Y
    )

    $tile             = New-Object System.Windows.Forms.Panel
    $tile.Size        = New-Object System.Drawing.Size(330, 340)
    $tile.Location    = New-Object System.Drawing.Point($X, $Y)
    $tile.BackColor   = $BrandBlue
    $tile.Cursor      = [System.Windows.Forms.Cursors]::Hand

    $headLbl          = New-Object System.Windows.Forms.Label
    $headLbl.Text     = $Category.ToUpper()
    $headLbl.Font     = New-Object System.Drawing.Font('Segoe UI', 19, [System.Drawing.FontStyle]::Bold)
    $headLbl.ForeColor= [System.Drawing.Color]::White
    $headLbl.BackColor= [System.Drawing.Color]::Transparent
    $headLbl.Location = New-Object System.Drawing.Point(0, 70)
    $headLbl.Size     = New-Object System.Drawing.Size(330, 40)
    $headLbl.TextAlign= 'MiddleCenter'
    $headLbl.Cursor   = [System.Windows.Forms.Cursors]::Hand
    $tile.Controls.Add($headLbl)

    $countLbl          = New-Object System.Windows.Forms.Label
    $countLbl.Text     = "$ModuleCount module(s)"
    $countLbl.Font     = New-Object System.Drawing.Font('Segoe UI', 11)
    $countLbl.ForeColor= [System.Drawing.Color]::White
    $countLbl.BackColor= [System.Drawing.Color]::Transparent
    $countLbl.Location = New-Object System.Drawing.Point(0, 115)
    $countLbl.Size     = New-Object System.Drawing.Size(330, 22)
    $countLbl.TextAlign= 'MiddleCenter'
    $countLbl.Cursor   = [System.Windows.Forms.Cursors]::Hand
    $tile.Controls.Add($countLbl)

    $taglineLbl          = New-Object System.Windows.Forms.Label
    $taglineLbl.Text     = $Tagline
    $taglineLbl.Font     = New-Object System.Drawing.Font('Segoe UI', 10)
    $taglineLbl.ForeColor= [System.Drawing.Color]::White
    $taglineLbl.BackColor= [System.Drawing.Color]::Transparent
    $taglineLbl.Location = New-Object System.Drawing.Point(20, 160)
    $taglineLbl.Size     = New-Object System.Drawing.Size(290, 90)
    $taglineLbl.TextAlign= 'MiddleCenter'
    $taglineLbl.Cursor   = [System.Windows.Forms.Cursors]::Hand
    $tile.Controls.Add($taglineLbl)

    $cueLbl          = New-Object System.Windows.Forms.Label
    $cueLbl.Text     = "Click to open >"
    $cueLbl.Font     = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $cueLbl.ForeColor= [System.Drawing.Color]::White
    $cueLbl.BackColor= [System.Drawing.Color]::Transparent
    $cueLbl.Location = New-Object System.Drawing.Point(0, 280)
    $cueLbl.Size     = New-Object System.Drawing.Size(330, 22)
    $cueLbl.TextAlign= 'MiddleCenter'
    $cueLbl.Cursor   = [System.Windows.Forms.Cursors]::Hand
    $tile.Controls.Add($cueLbl)

    # Wire click on the tile AND on every child label
    $clickHandler = { Show-Category -Category $Category }.GetNewClosure()
    $tile.Add_Click($clickHandler)
    foreach ($child in @($headLbl, $countLbl, $taglineLbl, $cueLbl)) {
        $child.Add_Click($clickHandler)
    }

    return $tile
}

# ----- prerequisite status banner (landing page) -----
$BrandGreen = [System.Drawing.ColorTranslator]::FromHtml('#28A745')
$BrandRed   = [System.Drawing.ColorTranslator]::FromHtml('#C0392B')

function New-StatusPill {
    param ([string]$Name, [bool]$Ok, [string]$Detail, [int]$X, [int]$Y)

    $pill              = New-Object System.Windows.Forms.Panel
    $pill.Size         = New-Object System.Drawing.Size(330, 36)
    $pill.Location     = New-Object System.Drawing.Point($X, $Y)
    $pill.BackColor    = if ($Ok) { $BrandGreen } else { $BrandRed }

    $tag               = New-Object System.Windows.Forms.Label
    $tag.Text          = if ($Ok) { 'OK' } else { 'MISSING' }
    $tag.Font          = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $tag.ForeColor     = [System.Drawing.Color]::White
    $tag.BackColor     = [System.Drawing.Color]::Transparent
    $tag.Location      = New-Object System.Drawing.Point(10, 8)
    $tag.Size          = New-Object System.Drawing.Size(70, 20)
    $tag.TextAlign     = 'MiddleLeft'
    $pill.Controls.Add($tag)

    $lbl               = New-Object System.Windows.Forms.Label
    $lbl.Text          = $Name
    $lbl.Font          = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $lbl.ForeColor     = [System.Drawing.Color]::White
    $lbl.BackColor     = [System.Drawing.Color]::Transparent
    $lbl.Location      = New-Object System.Drawing.Point(80, 4)
    $lbl.Size          = New-Object System.Drawing.Size(180, 18)
    $lbl.TextAlign     = 'MiddleLeft'
    $pill.Controls.Add($lbl)

    $det               = New-Object System.Windows.Forms.Label
    $det.Text          = $Detail
    $det.Font          = New-Object System.Drawing.Font('Segoe UI', 8)
    $det.ForeColor     = [System.Drawing.Color]::White
    $det.BackColor     = [System.Drawing.Color]::Transparent
    $det.Location      = New-Object System.Drawing.Point(80, 19)
    $det.Size          = New-Object System.Drawing.Size(240, 14)
    $det.TextAlign     = 'MiddleLeft'
    $pill.Controls.Add($det)

    return $pill
}

$prereqs = Get-RequiredPrereqStatus
# 3 pills, 330 wide each, 15px gap, centred in the 1065-wide category panel
# Total width: 3*330 + 2*15 = 1020; left margin = (1065-1020)/2 = 22 (round to 22)
$pillY = 10
$pillX = 22
foreach ($p in $prereqs) {
    $pill = New-StatusPill -Name $p.Name -Ok $p.Ok -Detail $p.Detail -X $pillX -Y $pillY
    $categoryPanel.Controls.Add($pill)
    $pillX += 345
}

$missingCount = @($prereqs | Where-Object { -not $_.Ok }).Count
if ($missingCount -gt 0) {
    $hintLbl           = New-Object System.Windows.Forms.Label
    $hintLbl.Text      = "$missingCount prerequisite(s) missing. Open the Utility tile to install."
    $hintLbl.Font      = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Italic)
    $hintLbl.ForeColor = $BrandRed
    $hintLbl.Location  = New-Object System.Drawing.Point(22, 52)
    $hintLbl.Size      = New-Object System.Drawing.Size(1020, 18)
    $hintLbl.TextAlign = 'MiddleCenter'
    $categoryPanel.Controls.Add($hintLbl)
}

# ----- category tiles (3 across) -----
$reportingCount = @($script:Modules | Where-Object { $_.Category -eq 'Reporting' }).Count
$adminCount     = @($script:Modules | Where-Object { $_.Category -eq 'Administration' }).Count
$utilityCount   = @($script:Modules | Where-Object { $_.Category -eq 'Utility' }).Count

# 3 tiles, 330 wide, 15px gap, ~22 left margin (same as pills row)
$tileY = 78
$reportingTile = New-CategoryTile `
    -Category    'Reporting' `
    -Tagline     "Tenant inventory reports: mailboxes, users, group membership, and MFA / authentication methods." `
    -ModuleCount $reportingCount `
    -X 22 -Y $tileY

$adminTile = New-CategoryTile `
    -Category    'Administration' `
    -Tagline     "Make changes in the tenant: distribution lists, contacts, guests, calendar and mailbox access." `
    -ModuleCount $adminCount `
    -X 367 -Y $tileY

$utilityTile = New-CategoryTile `
    -Category    'Utility' `
    -Tagline     "Install prerequisites: ExchangeOnlineManagement, Microsoft Graph modules, PowerShell 7." `
    -ModuleCount $utilityCount `
    -X 712 -Y $tileY

$categoryPanel.Controls.Add($reportingTile)
$categoryPanel.Controls.Add($adminTile)
$categoryPanel.Controls.Add($utilityTile)

# ----- footer -----
$statusLbl.Text      = 'Ready.'
$statusLbl.Location  = New-Object System.Drawing.Point(18, 538)
$statusLbl.Size      = New-Object System.Drawing.Size(680, 22)
$statusLbl.TextAlign = 'MiddleLeft'
$statusLbl.ForeColor = $BrandTextMuted
$statusLbl.Font      = New-Object System.Drawing.Font('Segoe UI', 9)
$form.Controls.Add($statusLbl)

$aboutBtn           = New-Object System.Windows.Forms.Button
$aboutBtn.Text      = 'About'
$aboutBtn.Size      = New-Object System.Drawing.Size(90, 32)
$aboutBtn.Location  = New-Object System.Drawing.Point(720, 533)
$aboutBtn.FlatStyle = 'Flat'
$aboutBtn.BackColor = [System.Drawing.Color]::White
$aboutBtn.ForeColor = $BrandTextDark
$aboutBtn.Add_Click({ Show-AboutDialog })
$form.Controls.Add($aboutBtn)

$openBtn           = New-Object System.Windows.Forms.Button
$openBtn.Text      = 'View Results'
$openBtn.Size      = New-Object System.Drawing.Size(150, 32)
$openBtn.Location  = New-Object System.Drawing.Point(820, 533)
$openBtn.FlatStyle = 'Flat'
$openBtn.BackColor = [System.Drawing.Color]::White
$openBtn.ForeColor = $BrandTextDark
$openBtn.Add_Click({ Open-ResultsDir })
$form.Controls.Add($openBtn)

$closeBtn           = New-Object System.Windows.Forms.Button
$closeBtn.Text      = 'Close'
$closeBtn.Size      = New-Object System.Drawing.Size(90, 32)
$closeBtn.Location  = New-Object System.Drawing.Point(980, 533)
$closeBtn.FlatStyle = 'Flat'
$closeBtn.BackColor = [System.Drawing.Color]::White
$closeBtn.ForeColor = $BrandTextDark
$closeBtn.Add_Click({ $form.Close() })
$form.Controls.Add($closeBtn)

[void]$form.ShowDialog()
$form.Dispose()
