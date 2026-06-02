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
$script:Version    = '0.4.0-beta'
$script:BaseRawUrl = "https://raw.githubusercontent.com/$script:RepoOwner/$script:RepoName/$script:Branch"
$script:WorkDir    = Join-Path $env:TEMP 'MSPM365Utility'   # module cache (internal)
$script:ResultsDir = 'C:\MSP-M365-Utility'                  # where reporting modules drop CSVs

# Module catalog. When a new module is added to /Modules, add it here too.
$script:Modules = @(
    [PSCustomObject]@{
        File        = 'Modules/Get-TenantMailboxes.ps1'
        Title       = 'Tenant Mailbox Inventory'
        Category    = 'Reporting'
        Description = 'Export every mailbox (user, shared, resource) with enabled and licensed status to CSV.'
    }
    [PSCustomObject]@{
        File        = 'Modules/Get-TenantUserDetails.ps1'
        Title       = 'Tenant User Details'
        Category    = 'Reporting'
        Description = 'Export user details - name, email, title, department, manager - to CSV.'
    }
    [PSCustomObject]@{
        File        = 'Modules/Get-TenantGroupMembership.ps1'
        Title       = 'Tenant Group Membership'
        Category    = 'Reporting'
        Description = 'Export all groups (DL, mail-enabled security, M365, security) and members to CSV.'
    }
    [PSCustomObject]@{
        File        = 'Modules/Get-TenantMFAStatus.ps1'
        Title       = 'Tenant MFA Status'
        Category    = 'Reporting'
        Description = 'Licensed users (guests excluded) with MFA status, default method, and Authenticator/SMS/OTP flags.'
    }
    [PSCustomObject]@{
        File        = 'Modules/Add-DistroMember.ps1'
        Title       = 'Add Distribution List Members'
        Category    = 'Administration'
        Description = 'Bulk add members to a distribution list from a pasted email list.'
    }
    [PSCustomObject]@{
        File        = 'Modules/Remove-DistroMember.ps1'
        Title       = 'Remove Distribution List Members'
        Category    = 'Administration'
        Description = 'Bulk remove members from a distribution list from a pasted email list.'
    }
    [PSCustomObject]@{
        File        = 'Modules/Add-BulkContact.ps1'
        Title       = 'Bulk Contact Creation'
        Category    = 'Administration'
        Description = 'Bulk create external Mail Contacts from a paste-list. Prompts per duplicate.'
    }
    [PSCustomObject]@{
        File        = 'Modules/Add-BulkGuestUser.ps1'
        Title       = 'Bulk Guest User Invite'
        Category    = 'Administration'
        Description = 'Bulk send B2B guest invitations. Auto-detects tenant redirect URL. Prompts per duplicate.'
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
        Start-Process -FilePath $psExe -ArgumentList @(
            '-NoExit',
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
$form.Size          = New-Object System.Drawing.Size(800, 620)
$form.AutoScaleMode = 'Dpi'
$form.StartPosition = 'CenterScreen'
$form.BackColor     = [System.Drawing.Color]::White
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox   = $false
$form.Font          = New-Object System.Drawing.Font('Segoe UI', 9)

# ----- header -----
$header           = New-Object System.Windows.Forms.Panel
$header.Size      = New-Object System.Drawing.Size(800, 72)
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
$subLbl.Text     = 'Each module opens in its own PowerShell window. Prompts and progress appear there.'
$subLbl.Font     = New-Object System.Drawing.Font('Segoe UI', 9)
$subLbl.ForeColor= [System.Drawing.Color]::White
$subLbl.Location = New-Object System.Drawing.Point(22, 42)
$subLbl.Size     = New-Object System.Drawing.Size(700, 20)
$subLbl.BackColor= [System.Drawing.Color]::Transparent
$header.Controls.Add($subLbl)

# ----- module list -----
$list                 = New-Object System.Windows.Forms.FlowLayoutPanel
$list.Location        = New-Object System.Drawing.Point(15, 85)
$list.Size            = New-Object System.Drawing.Size(760, 430)
$list.FlowDirection   = 'TopDown'
$list.WrapContents    = $false
$list.AutoScroll      = $true
$list.BackColor       = [System.Drawing.Color]::White
$form.Controls.Add($list)

# Forward-declare so card click handlers can reach it
$statusLbl = New-Object System.Windows.Forms.Label

foreach ($mod in $script:Modules) {
    $card             = New-Object System.Windows.Forms.Panel
    $card.Size        = New-Object System.Drawing.Size(735, 76)
    $card.Margin      = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
    $card.BackColor   = $BrandBlueLight
    $card.BorderStyle = 'FixedSingle'

    $catLbl           = New-Object System.Windows.Forms.Label
    $catLbl.Text      = $mod.Category.ToUpper()
    $catLbl.Font      = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)
    $catLbl.ForeColor = $BrandBlueDark
    $catLbl.Location  = New-Object System.Drawing.Point(12, 8)
    $catLbl.Size      = New-Object System.Drawing.Size(120, 14)
    $catLbl.BackColor = [System.Drawing.Color]::Transparent
    $card.Controls.Add($catLbl)

    $modTitleLbl          = New-Object System.Windows.Forms.Label
    $modTitleLbl.Text     = $mod.Title
    $modTitleLbl.Font     = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
    $modTitleLbl.ForeColor= $BrandTextDark
    $modTitleLbl.Location = New-Object System.Drawing.Point(12, 22)
    $modTitleLbl.Size     = New-Object System.Drawing.Size(540, 22)
    $modTitleLbl.BackColor= [System.Drawing.Color]::Transparent
    $card.Controls.Add($modTitleLbl)

    $descLbl          = New-Object System.Windows.Forms.Label
    $descLbl.Text     = $mod.Description
    $descLbl.Font     = New-Object System.Drawing.Font('Segoe UI', 9)
    $descLbl.ForeColor= $BrandTextDark
    $descLbl.Location = New-Object System.Drawing.Point(12, 46)
    $descLbl.Size     = New-Object System.Drawing.Size(540, 24)
    $descLbl.BackColor= [System.Drawing.Color]::Transparent
    $card.Controls.Add($descLbl)

    $runBtn           = New-Object System.Windows.Forms.Button
    $runBtn.Text      = 'Run'
    $runBtn.Size      = New-Object System.Drawing.Size(140, 40)
    $runBtn.Location  = New-Object System.Drawing.Point(575, 18)
    $runBtn.BackColor = $BrandBlue
    $runBtn.ForeColor = [System.Drawing.Color]::White
    $runBtn.FlatStyle = 'Flat'
    $runBtn.FlatAppearance.BorderSize = 0
    $runBtn.Font      = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
    $runBtn.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $localMod         = $mod
    $localStatus      = $statusLbl
    $runBtn.Add_Click({ Invoke-Module -Module $localMod -StatusLabel $localStatus }.GetNewClosure())
    $card.Controls.Add($runBtn)

    $list.Controls.Add($card)
}

# ----- footer -----
$statusLbl.Text      = 'Ready.'
$statusLbl.Location  = New-Object System.Drawing.Point(18, 538)
$statusLbl.Size      = New-Object System.Drawing.Size(380, 22)
$statusLbl.TextAlign = 'MiddleLeft'
$statusLbl.ForeColor = $BrandTextMuted
$statusLbl.Font      = New-Object System.Drawing.Font('Segoe UI', 9)
$form.Controls.Add($statusLbl)

$aboutBtn           = New-Object System.Windows.Forms.Button
$aboutBtn.Text      = 'About'
$aboutBtn.Size      = New-Object System.Drawing.Size(90, 32)
$aboutBtn.Location  = New-Object System.Drawing.Point(420, 533)
$aboutBtn.FlatStyle = 'Flat'
$aboutBtn.BackColor = [System.Drawing.Color]::White
$aboutBtn.ForeColor = $BrandTextDark
$aboutBtn.Add_Click({ Show-AboutDialog })
$form.Controls.Add($aboutBtn)

$openBtn           = New-Object System.Windows.Forms.Button
$openBtn.Text      = 'View Results'
$openBtn.Size      = New-Object System.Drawing.Size(150, 32)
$openBtn.Location  = New-Object System.Drawing.Point(520, 533)
$openBtn.FlatStyle = 'Flat'
$openBtn.BackColor = [System.Drawing.Color]::White
$openBtn.ForeColor = $BrandTextDark
$openBtn.Add_Click({ Open-ResultsDir })
$form.Controls.Add($openBtn)

$closeBtn           = New-Object System.Windows.Forms.Button
$closeBtn.Text      = 'Close'
$closeBtn.Size      = New-Object System.Drawing.Size(90, 32)
$closeBtn.Location  = New-Object System.Drawing.Point(680, 533)
$closeBtn.FlatStyle = 'Flat'
$closeBtn.BackColor = [System.Drawing.Color]::White
$closeBtn.ForeColor = $BrandTextDark
$closeBtn.Add_Click({ $form.Close() })
$form.Controls.Add($closeBtn)

[void]$form.ShowDialog()
$form.Dispose()
