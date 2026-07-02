#Requires -Version 5.1

<#
.SYNOPSIS
    Install or update the Microsoft Graph PowerShell submodules used by the
    MSP M365 Utility scripts.

.DESCRIPTION
    Installs the specific Microsoft.Graph.* submodules used by the reporting
    and administration modules:
        - Microsoft.Graph.Users
        - Microsoft.Graph.Users.Actions
        - Microsoft.Graph.Groups
        - Microsoft.Graph.Reports
        - Microsoft.Graph.Identity.SignIns

    The full Microsoft.Graph umbrella module is ~150MB and pulls in 30+
    submodules. Installing only what we need is much faster and uses much
    less disk space.

.NOTES
    Does NOT require admin (uses -Scope CurrentUser).
#>

$ErrorActionPreference = 'Stop'
$Host.UI.RawUI.WindowTitle = "Install Microsoft Graph Modules"

Clear-Host
Write-Host ""
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |     Install Microsoft Graph PS Submodules        |" -ForegroundColor Cyan
Write-Host "  +--------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""

$GraphModules = @(
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Users.Actions',
    'Microsoft.Graph.Groups',
    'Microsoft.Graph.Reports',
    'Microsoft.Graph.Identity.SignIns'
)

# Show current state
$status = foreach ($m in $GraphModules) {
    $installed = Get-Module -ListAvailable -Name $m -ErrorAction SilentlyContinue
    [PSCustomObject]@{
        Name      = $m
        Installed = [bool]$installed
        Version   = if ($installed) { ($installed | Sort-Object Version -Descending | Select-Object -First 1).Version.ToString() } else { '-' }
    }
}
Write-Host "  Current state:" -ForegroundColor Cyan
$status | ForEach-Object {
    $tag = if ($_.Installed) { '[OK]    ' } else { '[MISS]  ' }
    $col = if ($_.Installed) { 'Green' } else { 'Yellow' }
    Write-Host ("  $tag {0,-40} {1}" -f $_.Name, $_.Version) -ForegroundColor $col
}
Write-Host ""

$missingCount = ($status | Where-Object { -not $_.Installed }).Count
if ($missingCount -eq 0) {
    $answer = Read-Host "  All present. Reinstall / update to latest? (Y/N) [N]"
    if ($answer -notin @('Y','y')) {
        Write-Host "  Nothing to do." -ForegroundColor DarkGray
        Read-Host "`n  Press Enter to exit"; exit 0
    }
}

# Install
Write-Host ""
Write-Host "  Installing $($GraphModules.Count) module(s) to CurrentUser scope..." -ForegroundColor Cyan
$failed = @()
foreach ($m in $GraphModules) {
    Write-Host "  -> $m" -ForegroundColor DarkCyan
    try {
        Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        $v = (Get-Module -ListAvailable -Name $m | Sort-Object Version -Descending | Select-Object -First 1).Version
        Write-Host "     [OK] v$v" -ForegroundColor Green
    } catch {
        Write-Host "     [ERROR] $($_.Exception.Message)" -ForegroundColor Red
        $failed += $m
    }
}

Write-Host ""
if ($failed.Count -eq 0) {
    Write-Host "  All Graph submodules installed successfully." -ForegroundColor Green
} else {
    Write-Host "  Failed: $($failed -join ', ')" -ForegroundColor Red
}
Read-Host "`n  Press Enter to exit"
