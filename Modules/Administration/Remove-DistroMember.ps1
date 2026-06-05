# Prompt for Distribution List
$DL = Read-Host "Enter the Distribution List email address"

# Prompt for members - paste all emails at once
Write-Host "`nPaste all member emails to REMOVE below (one per line)." -ForegroundColor Yellow
Write-Host "When done, press ENTER on a blank line to continue:`n" -ForegroundColor Yellow

$Members = @()
while ($true) {
    $line = Read-Host
    if ($line -eq "") { break }
    $Members += $line.Trim()
}

# Clean and parse the list (handles commas, semicolons, extra spaces, blank lines)
$Members = $Members |
    ForEach-Object { $_ -split '[,;\s]+' } |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ -match '^[^@\s]+@[^@\s]+\.[^@\s]+$' }  # basic email validation

# Preview before executing
Write-Host "`nDistribution List  : $DL" -ForegroundColor Cyan
Write-Host "Members to REMOVE  : $($Members.Count)" -ForegroundColor Cyan
$Members | ForEach-Object { Write-Host "  - $_" }

$Confirm = Read-Host "`nProceed with REMOVAL? (Y/N)"
if ($Confirm -notin @("Y","y")) {
    Write-Host "Aborted." -ForegroundColor Red
    Read-Host "`nPress Enter to exit"
    exit
}

# Connect to Exchange Online
Connect-ExchangeOnline

# Remove members
foreach ($Member in $Members) {
    try {
        Remove-DistributionGroupMember -Identity $DL -Member $Member -Confirm:$false -ErrorAction Stop
        Write-Host "Removed : $Member" -ForegroundColor Green
    } catch {
        Write-Host "Failed  : $Member — $_" -ForegroundColor Red
    }
}

Write-Host "`nDone!" -ForegroundColor Cyan
Read-Host "`nPress Enter to exit"