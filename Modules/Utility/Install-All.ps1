#Requires -Version 5.1

<#
.SYNOPSIS
    Install everything the MSP M365 Utility needs:
        - ExchangeOnlineManagement PowerShell module
        - Microsoft.Graph.* submodules (Users, Groups, Reports, Identity.SignIns)
        - PowerShell 7 (via winget)

.DESCRIPTION
    Runs all three installs sequentially. Each step is idempotent and skips
    re-install if already present (no prompt - this is the "do everything"
    convenience script).

.NOTES
    Does NOT require admin (Install-Module -Scope CurrentUser; winget runs
    as the current user).
#>

$ErrorActionPreference = 'Stop'
$Host.UI.RawUI.WindowTitle = "Install All Prerequisites"

Clear-Host
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |          Install All Prerequisites               |" -ForegroundColor Cyan
Write-Host "  |     ExchangeOnline + MS Graph + PowerShell 7     |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""

$StepResults = [System.Collections.Generic.List[PSCustomObject]]::new()

# ----------------------------------------------------------
#  1) Exchange Online Management
# ----------------------------------------------------------
Write-Host "  [1/3] ExchangeOnlineManagement" -ForegroundColor Cyan
try {
    $existing = Get-Module -ListAvailable -Name 'ExchangeOnlineManagement' -ErrorAction SilentlyContinue
    if ($existing) {
        $v = ($existing | Sort-Object Version -Descending | Select-Object -First 1).Version
        Write-Host "    Already installed (v$v) - skipping." -ForegroundColor DarkGray
        $StepResults.Add([PSCustomObject]@{ Step='ExchangeOnlineManagement'; Status='AlreadyInstalled'; Detail="v$v" })
    } else {
        Install-Module -Name 'ExchangeOnlineManagement' -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        $v = (Get-Module -ListAvailable -Name 'ExchangeOnlineManagement' | Sort-Object Version -Descending | Select-Object -First 1).Version
        Write-Host "    [OK] Installed v$v" -ForegroundColor Green
        $StepResults.Add([PSCustomObject]@{ Step='ExchangeOnlineManagement'; Status='Installed'; Detail="v$v" })
    }
} catch {
    Write-Host "    [ERROR] $($_.Exception.Message)" -ForegroundColor Red
    $StepResults.Add([PSCustomObject]@{ Step='ExchangeOnlineManagement'; Status='Failed'; Detail=$_.Exception.Message })
}

# ----------------------------------------------------------
#  2) Microsoft Graph submodules
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [2/3] Microsoft Graph submodules" -ForegroundColor Cyan
$GraphModules = @(
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Groups',
    'Microsoft.Graph.Reports',
    'Microsoft.Graph.Identity.SignIns'
)
foreach ($m in $GraphModules) {
    try {
        $existing = Get-Module -ListAvailable -Name $m -ErrorAction SilentlyContinue
        if ($existing) {
            $v = ($existing | Sort-Object Version -Descending | Select-Object -First 1).Version
            Write-Host "    $m  -  already installed (v$v) - skipping." -ForegroundColor DarkGray
            $StepResults.Add([PSCustomObject]@{ Step=$m; Status='AlreadyInstalled'; Detail="v$v" })
        } else {
            Write-Host "    Installing $m ..." -ForegroundColor DarkCyan
            Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            $v = (Get-Module -ListAvailable -Name $m | Sort-Object Version -Descending | Select-Object -First 1).Version
            Write-Host "       [OK] v$v" -ForegroundColor Green
            $StepResults.Add([PSCustomObject]@{ Step=$m; Status='Installed'; Detail="v$v" })
        }
    } catch {
        Write-Host "       [ERROR] $($_.Exception.Message)" -ForegroundColor Red
        $StepResults.Add([PSCustomObject]@{ Step=$m; Status='Failed'; Detail=$_.Exception.Message })
    }
}

# ----------------------------------------------------------
#  3) PowerShell 7 via winget
# ----------------------------------------------------------
Write-Host ""
Write-Host "  [3/3] PowerShell 7 (winget)" -ForegroundColor Cyan
$existing = Get-Command pwsh -ErrorAction SilentlyContinue
if ($existing) {
    try {
        $ver = & pwsh -NoProfile -Command '$PSVersionTable.PSVersion.ToString()'
        Write-Host "    Already installed (v$ver) - skipping." -ForegroundColor DarkGray
        $StepResults.Add([PSCustomObject]@{ Step='PowerShell 7'; Status='AlreadyInstalled'; Detail="v$ver" })
    } catch {
        Write-Host "    pwsh present (version probe failed) - skipping." -ForegroundColor DarkGray
        $StepResults.Add([PSCustomObject]@{ Step='PowerShell 7'; Status='AlreadyInstalled'; Detail=$existing.Source })
    }
} else {
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $winget) {
        Write-Host "    [ERROR] winget not found - cannot install PowerShell 7 automatically." -ForegroundColor Red
        Write-Host "    Install 'App Installer' from the Microsoft Store, then re-run this." -ForegroundColor Yellow
        $StepResults.Add([PSCustomObject]@{ Step='PowerShell 7'; Status='Failed'; Detail='winget not available' })
    } else {
        try {
            & winget install --id Microsoft.PowerShell --source winget `
                --accept-package-agreements --accept-source-agreements --silent
            $exit = $LASTEXITCODE
            $check = Get-Command pwsh -ErrorAction SilentlyContinue
            if (-not $check) {
                $candidate = Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'
                if (Test-Path $candidate) {
                    Write-Host "    [OK] Installed (open a new terminal so PATH refreshes)." -ForegroundColor Green
                    $StepResults.Add([PSCustomObject]@{ Step='PowerShell 7'; Status='Installed'; Detail='installed; PATH refresh needed' })
                } else {
                    Write-Host "    [WARN] winget exit $exit; pwsh not yet visible." -ForegroundColor Yellow
                    $StepResults.Add([PSCustomObject]@{ Step='PowerShell 7'; Status='Unknown'; Detail="winget exit $exit" })
                }
            } else {
                Write-Host "    [OK] PowerShell 7 installed." -ForegroundColor Green
                $StepResults.Add([PSCustomObject]@{ Step='PowerShell 7'; Status='Installed'; Detail=$check.Source })
            }
        } catch {
            Write-Host "    [ERROR] $_" -ForegroundColor Red
            $StepResults.Add([PSCustomObject]@{ Step='PowerShell 7'; Status='Failed'; Detail=$_.Exception.Message })
        }
    }
}

# ----------------------------------------------------------
#  SUMMARY
# ----------------------------------------------------------
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |                    SUMMARY                       |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
foreach ($r in $StepResults) {
    $col = switch ($r.Status) {
        'Installed'        { 'Green' }
        'AlreadyInstalled' { 'DarkGray' }
        'Failed'           { 'Red' }
        default            { 'Yellow' }
    }
    Write-Host ("  {0,-32} {1,-18} {2}" -f $r.Step, $r.Status, $r.Detail) -ForegroundColor $col
}
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""
Read-Host "  Press Enter to exit"
