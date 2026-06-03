# MSP M365 Utility

PowerShell modules for Microsoft 365 tenant administration and reporting in an MSP context.

## Quick Start (GUI Launcher)

Run on any Windows machine — no clone needed:

```powershell
iex (irm "https://raw.githubusercontent.com/MasatoNakajima20/MSP-M365-Utility/main/Launch-MSPM365Utility.ps1")
```

The launcher fetches each module on demand and opens it in its own PowerShell window so interactive prompts (tenant code, distribution list, paste-lists) work normally.

- **Module cache:** `%TEMP%\MSPM365Utility\` (downloaded `.ps1` files)
- **Report output:** `C:\MSP-M365-Utility\` (all CSV results)

Use the **View Results** button in the launcher to jump straight to the report folder.

## Modules

Modules are organised by category. The launcher's landing screen lets you pick a category, then shows that category's modules.

### Reporting (`Modules/Reporting/`)

| Script | Purpose |
| --- | --- |
| `Get-TenantUserDetails.ps1` | Pull user details for a tenant |
| `Get-TenantMailboxes.ps1` | Enumerate mailboxes in a tenant |
| `Get-TenantGroupMembership.ps1` | Report group membership across a tenant |
| `Get-TenantMFAStatus.ps1` | MFA status and authentication methods for licensed users (guests excluded) |

### Administration (`Modules/Administration/`)

| Script | Purpose |
| --- | --- |
| `Add-DistroMember.ps1` | Add a member to a distribution group |
| `Remove-DistroMember.ps1` | Remove a member from a distribution group |
| `Add-BulkContact.ps1` | Bulk-create external Mail Contacts from a pasted list |
| `Add-BulkGuestUser.ps1` | Bulk-send B2B guest invitations (auto-detects tenant redirect URL) |
| `Add-CalendarPermission.ps1` | Grant Reviewer / Author / Editor access on a mailbox's calendar (per-user role) |
| `Remove-CalendarPermission.ps1` | Remove user(s) calendar access on a target mailbox |
| `Add-MailboxAccess.ps1` | Bulk grant FullAccess + SendAs (creates target as shared mailbox if missing) |
| `Remove-MailboxAccess.ps1` | Bulk remove FullAccess + SendAs from a target mailbox |

## Requirements

- PowerShell 7+
- Microsoft Graph PowerShell SDK and/or ExchangeOnlineManagement modules (as required per script)
- Appropriate delegated/admin permissions in the target tenant

## Usage

Import or dot-source the script you need, then call the function with the required parameters. See the comment-based help inside each script for parameter details.
