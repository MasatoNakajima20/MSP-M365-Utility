# Changelog

All notable changes to **MSP M365 Utility** are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/); this
project uses SemVer with a `-beta` suffix. This file was introduced at
`0.13.0-beta`; earlier version history is available via the git tags
(`0.1.0-beta` … `0.12.1-beta`).

## [0.13.0-beta] - 2026-09-21
### Added
- `Invoke-UserOffboarding`: hide the mailbox from the Global Address List
  (`Set-Mailbox -HiddenFromAddressListsEnabled $true`), applied always-on as
  step 6c (right after the convert-to-shared / forwarding / delegate steps),
  logged like every other action.
