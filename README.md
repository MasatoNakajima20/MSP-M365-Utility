# MSP M365 Utility

PowerShell modules for Microsoft 365 tenant administration and reporting in an MSP context.

## Modules

| Script | Purpose |
| --- | --- |
| `Modules/Get-TenantUserDetails.ps1` | Pull user details for a tenant |
| `Modules/Get-TenantMailboxes.ps1` | Enumerate mailboxes in a tenant |
| `Modules/Get-TenantGroupMembership.ps1` | Report group membership across a tenant |
| `Modules/Add-DistroMember.ps1` | Add a member to a distribution group |
| `Modules/Remove-DistroMember.ps1` | Remove a member from a distribution group |

## Requirements

- PowerShell 7+
- Microsoft Graph PowerShell SDK and/or ExchangeOnlineManagement modules (as required per script)
- Appropriate delegated/admin permissions in the target tenant

## Usage

Import or dot-source the script you need, then call the function with the required parameters. See the comment-based help inside each script for parameter details.
