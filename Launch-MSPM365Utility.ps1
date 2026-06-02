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
$script:Version    = '0.2.0-beta'
$script:BaseRawUrl = "https://raw.githubusercontent.com/$script:RepoOwner/$script:RepoName/$script:Branch"
$script:WorkDir    = Join-Path $env:TEMP 'MSPM365Utility'

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
        $StatusLabel.Text = "Launched: $($Module.Title)  (output lands in $script:WorkDir)"
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

# ----- form -----
$form               = New-Object System.Windows.Forms.Form
$form.Text          = "MSP M365 Utility  -  $script:Version"
$form.Size          = New-Object System.Drawing.Size(800, 600)
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
$statusLbl.Location  = New-Object System.Drawing.Point(15, 525)
$statusLbl.Size      = New-Object System.Drawing.Size(540, 18)
$statusLbl.ForeColor = $BrandTextMuted
$statusLbl.Font      = New-Object System.Drawing.Font('Segoe UI', 8)
$form.Controls.Add($statusLbl)

$openBtn           = New-Object System.Windows.Forms.Button
$openBtn.Text      = 'Open Output Folder'
$openBtn.Size      = New-Object System.Drawing.Size(150, 28)
$openBtn.Location  = New-Object System.Drawing.Point(465, 520)
$openBtn.FlatStyle = 'Flat'
$openBtn.BackColor = [System.Drawing.Color]::White
$openBtn.ForeColor = $BrandTextDark
$openBtn.Add_Click({ Open-WorkDir })
$form.Controls.Add($openBtn)

$closeBtn           = New-Object System.Windows.Forms.Button
$closeBtn.Text      = 'Close'
$closeBtn.Size      = New-Object System.Drawing.Size(140, 28)
$closeBtn.Location  = New-Object System.Drawing.Point(625, 520)
$closeBtn.FlatStyle = 'Flat'
$closeBtn.BackColor = [System.Drawing.Color]::White
$closeBtn.ForeColor = $BrandTextDark
$closeBtn.Add_Click({ $form.Close() })
$form.Controls.Add($closeBtn)

$repoLbl           = New-Object System.Windows.Forms.Label
$repoLbl.Text      = "github.com/$script:RepoOwner/$script:RepoName"
$repoLbl.Location  = New-Object System.Drawing.Point(15, 547)
$repoLbl.Size      = New-Object System.Drawing.Size(540, 16)
$repoLbl.ForeColor = $BrandTextMuted
$repoLbl.Font      = New-Object System.Drawing.Font('Segoe UI', 8)
$form.Controls.Add($repoLbl)

[void]$form.ShowDialog()
$form.Dispose()
