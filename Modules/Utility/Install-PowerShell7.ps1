#Requires -Version 5.1

<#
.SYNOPSIS
    Install PowerShell 7 via winget.

.DESCRIPTION
    Checks for an existing pwsh.exe on PATH. If absent, installs the
    Microsoft.PowerShell package via winget. Accepts package + source
    agreements non-interactively.

.NOTES
    Requires winget to be available (Windows 10 1809+ with App Installer,
    or any modern Windows 11). Does NOT require admin if winget is set up
    for the current user.
#>

$ErrorActionPreference = 'Stop'
$Host.UI.RawUI.WindowTitle = "Install PowerShell 7"

Clear-Host
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |             Install PowerShell 7                 |" -ForegroundColor Cyan
Write-Host "  |                  via winget                      |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""

# Check existing
$existing = Get-Command pwsh -ErrorAction SilentlyContinue
if ($existing) {
    try {
        $ver = & pwsh -NoProfile -Command '$PSVersionTable.PSVersion.ToString()'
        Write-Host "  PowerShell 7 already installed: v$ver" -ForegroundColor Green
        Write-Host "  Path: $($existing.Source)" -ForegroundColor DarkGray
    } catch {
        Write-Host "  pwsh found on PATH (version probe failed): $($existing.Source)" -ForegroundColor Green
    }
    $answer = Read-Host "`n  Reinstall / update to latest? (Y/N) [N]"
    if ($answer -notin @('Y','y')) {
        Write-Host "  Nothing to do." -ForegroundColor DarkGray
        Read-Host "`n  Press Enter to exit"; exit 0
    }
}

# Check winget
$winget = Get-Command winget -ErrorAction SilentlyContinue
if (-not $winget) {
    Write-Host "  [ERROR] winget not found on PATH." -ForegroundColor Red
    Write-Host "  Install 'App Installer' from the Microsoft Store, then try again." -ForegroundColor Yellow
    Read-Host "`n  Press Enter to exit"; exit 1
}

Write-Host ""
Write-Host "  Installing Microsoft.PowerShell via winget..." -ForegroundColor Cyan
try {
    & winget install --id Microsoft.PowerShell --source winget `
        --accept-package-agreements --accept-source-agreements --silent
    $exit = $LASTEXITCODE
    if ($exit -ne 0) {
        Write-Host "  [WARN] winget returned exit code $exit." -ForegroundColor Yellow
    }
} catch {
    Write-Host "  [ERROR] winget invocation failed: $_" -ForegroundColor Red
    Read-Host "`n  Press Enter to exit"; exit 1
}

# Verify - winget may have updated PATH for new shells; check both PATH and the default install path
Write-Host ""
Write-Host "  Verifying..." -ForegroundColor Cyan
$check = Get-Command pwsh -ErrorAction SilentlyContinue
if (-not $check) {
    $candidate = Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'
    if (Test-Path $candidate) {
        Write-Host "  [OK] Found pwsh at: $candidate" -ForegroundColor Green
        Write-Host "  (Open a new terminal so PATH picks up the install.)" -ForegroundColor Yellow
    } else {
        Write-Host "  [WARN] pwsh not yet visible. winget may still be finishing the install." -ForegroundColor Yellow
        Write-Host "  Open a new terminal and run:  pwsh -v" -ForegroundColor Yellow
    }
} else {
    try {
        $ver = & pwsh -NoProfile -Command '$PSVersionTable.PSVersion.ToString()'
        Write-Host "  [OK] PowerShell v$ver" -ForegroundColor Green
        Write-Host "  Path: $($check.Source)" -ForegroundColor DarkGray
    } catch {
        Write-Host "  [OK] pwsh on PATH: $($check.Source)" -ForegroundColor Green
    }
}

Read-Host "`n  Press Enter to exit"
