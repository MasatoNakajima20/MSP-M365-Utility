# Prompt for Distribution List
$DL = Read-Host "Enter the Distribution List email address"

# Prompt for members - paste all emails at once
Write-Host "`nPaste all member emails below (one per line)." -ForegroundColor Yellow
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
Write-Host "`nDistribution List : $DL" -ForegroundColor Cyan
Write-Host "Members to add    : $($Members.Count)" -ForegroundColor Cyan
$Members | ForEach-Object { Write-Host "  - $_" }

$Confirm = Read-Host "`nProceed? (Y/N)"
if ($Confirm -notin @("Y","y")) {
    Write-Host "Aborted." -ForegroundColor Red
    exit
}

# Connect to Exchange Online
Connect-ExchangeOnline

# Add members
foreach ($Member in $Members) {
    try {
        Add-DistributionGroupMember -Identity $DL -Member $Member -ErrorAction Stop
        Write-Host "Added  : $Member" -ForegroundColor Green
    } catch {
        Write-Host "Failed : $Member — $_" -ForegroundColor Red
    }
}

Write-Host "`nDone!" -ForegroundColor Cyan