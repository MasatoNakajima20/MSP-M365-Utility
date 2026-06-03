#Requires -Version 5.1

<#
.SYNOPSIS
    Install or update the ExchangeOnlineManagement PowerShell module.

.DESCRIPTION
    Installs ExchangeOnlineManagement to CurrentUser scope. If it's already
    present, prompts whether to reinstall / update.

.NOTES
    Does NOT require admin (uses -Scope CurrentUser).
#>

$ErrorActionPreference = 'Stop'
$Host.UI.RawUI.WindowTitle = "Install Exchange Online Module"

Clear-Host
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |       Install Exchange Online PS Module          |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""

$ModuleName = 'ExchangeOnlineManagement'

$installed = Get-Module -ListAvailable -Name $ModuleName -ErrorAction SilentlyContinue
if ($installed) {
    $latest = ($installed | Sort-Object Version -Descending | Select-Object -First 1).Version
    Write-Host "  $ModuleName is already installed (v$latest)." -ForegroundColor Green
    $answer = Read-Host "  Reinstall / update to latest? (Y/N) [N]"
    if ($answer -notin @('Y','y')) {
        Write-Host "  Nothing to do." -ForegroundColor DarkGray
        Read-Host "`n  Press Enter to exit"; exit 0
    }
}

Write-Host ""
Write-Host "  Installing $ModuleName (CurrentUser scope)..." -ForegroundColor Cyan
try {
    Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    $verified = (Get-Module -ListAvailable -Name $ModuleName | Sort-Object Version -Descending | Select-Object -First 1).Version
    Write-Host "  [OK] Installed $ModuleName v$verified." -ForegroundColor Green
} catch {
    Write-Host "  [ERROR] Install failed: $_" -ForegroundColor Red
    Read-Host "`n  Press Enter to exit"; exit 1
}

Write-Host ""
Write-Host "  Done." -ForegroundColor Cyan
Read-Host "`n  Press Enter to exit"
