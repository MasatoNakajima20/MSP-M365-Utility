# Changelog

All notable changes to **MSP M365 Utility** are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/); this
project uses SemVer with a `-beta` suffix. This file was introduced at
`0.13.0-beta`; earlier version history is available via the git tags
(`0.1.0-beta` … `0.12.1-beta`).

## [0.15.0-beta] - 2026-10-06
### Fixed
- `Invoke-UserOffboarding`: CRITICAL license-safety fix - licenses are now
  removed only when safe. The shared-mailbox conversion is verified via
  `RecipientTypeDetails`; a failed/unverified conversion, an unknown mailbox
  size, or a mailbox over the 50 GB cap all SKIP license removal and flag it
  RETAINED. "No mailbox" is treated as safe to remove; "already shared" as
  success. (Previously a failed conversion could strip every license and lead
  to mailbox deletion after the grace period.)
- `Invoke-UserOffboarding`: silent failures no longer report as Success - the
  SendAs lookup, the shared-mailbox FullAccess index build, the SKU map, and
  mailbox-statistics lookups are now recorded (WARN / Partial / Failed).
- `Invoke-UserOffboarding`: `-BypassSecurityGroupManagerCheck` added to
  `Remove-DistributionGroupMember` so mail-enabled security group removals do
  not fail when the operator is not the group manager.
### Changed
- `Invoke-UserOffboarding`: shared-mailbox FullAccess is now resolved from an
  index built once before the user loop instead of an O(users x mailboxes)
  scan per user.
- `Invoke-UserOffboarding`: the run summary reports Submitted / Processed /
  Skipped / Not Found separately.
- `Invoke-UserOffboarding`: the "Remove Group Memberships" and "Remove Shared
  Mailbox Access" rows now log the actual group names (tagged DL /
  Mail-Enabled Security / M365 / Security) and shared-mailbox addresses
  (SendAs + FullAccess), comma-separated, in the results CSV Detail column.
### Added
- `Invoke-UserOffboarding`: `Write-Log` function (INFO/WARN/ERROR, timestamped,
  console + file). A run log is written to
  `C:\Logging\MSP-M365-Utility\COMPUTER_<timestamp>_UserOffboarding.log`; the
  results CSV stays in `C:\MSP-M365-Utility\` and is renamed to
  `COMPUTER_<timestamp>_UserOffboarding_<Tenant>.csv`.

## [0.14.0-beta] - 2026-09-25
### Added
- `Get-TenantUserDetails`: two new columns - `Licenses` (assigned SKUs
  resolved to friendly product names via a `$SkuFriendlyNames` lookup table,
  falling back to the raw SkuPartNumber for unmapped SKUs) and `Roles`
  (directory role memberships). Both are semicolon-separated and de-duped.
- `Get-TenantUserDetails`: `With License` and `With Role` totals in the run
  summary.
### Changed
- `Get-TenantUserDetails`: added the `Directory.Read.All` Microsoft Graph scope
  (required for the license and role lookups).

## [0.13.0-beta] - 2026-09-21
### Added
- `Invoke-UserOffboarding`: hide the mailbox from the Global Address List
  (`Set-Mailbox -HiddenFromAddressListsEnabled $true`), applied always-on as
  step 6c (right after the convert-to-shared / forwarding / delegate steps),
  logged like every other action.
