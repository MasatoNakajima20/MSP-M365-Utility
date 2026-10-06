# PROJECT_STATE.md — MSP M365 Utility

Living status log for the MSP M365 Utility. Update after every meaningful change.

- **Repo:** https://github.com/MasatoNakajima20/MSP-M365-Utility (public)
- **Local path:** `C:\Claude Projects\MSP 365 Reporting Tool`
- **Current release:** `0.15.1-beta`
- **Launcher `$script:Version`:** `0.15.1-beta`
- **Last updated:** 2026-10-06

---

## Versioning rule (standing)

As of 0.13.0-beta this project follows the **CLAUDE.md scheme** (not the older
patch/minor-or-500-lines rule):

- **Patch (0.0.x)** — bugfix / patch only.
- **Minor (0.x.0)** — any new, updated, or removed function.
- **Major (x.0.0)** — breaking/backward-incompatible change or major overhaul.
- Beta suffix stays until told otherwise.

Git workflow (updated 2026-09-25): push to a branch named after the version
number **without** a `v` prefix (e.g. `0.14.0`), one branch per version, as
directed by the operator. **No git tags** - the operator asked to stop tagging
releases (2026-09-25). Never add a Claude co-author line. Ask for the commit
message before every commit. On each version: set `$script:Version`, append a
CHANGELOG entry, and update this file.

Earlier history: 0.1.0-beta .. 0.13.0-beta were pushed directly to `main` and
tagged; that direct-to-main + tag workflow was retired at 0.14.0-beta.

---

## Done

### Core
- Single-file WinForms launcher (`Launch-MSPM365Utility.ps1`), runnable via
  `iex (irm <raw main URL>)`; fetches each module from GitHub on demand into
  `%TEMP%\MSPM365Utility\` and runs it in its own PowerShell window.
- Landing page: category tiles (Reporting / Administration / Utility) plus a
  prerequisite status banner (Exchange Online, MS Graph [5 submodules],
  PowerShell 7) with green/red pills. Version-aware version detection across
  all module scopes on disk.
- All report/results output standardized to `C:\MSP-M365-Utility\`.
- Module windows close on completion; every exit path pauses ("Press Enter to
  exit") so output stays readable.
- EXO + Graph modules connect **Graph first** to avoid the MSAL "WithLogging"
  assembly conflict.

### Reporting modules (`Modules/Reporting/`)
- `Get-TenantMailboxes` — mailbox inventory; type, enabled, licensed, LastSignIn,
  Stale (>90 days).
- `Get-TenantUserDetails` — user details with scope prompt (Members/All), incl.
  City/State/Country/phones, AccountStatus, Licenses (friendly SKU names) and
  Roles (directory roles).
- `Get-TenantGroupMembership` — all group types and members.
- `Get-TenantMFAStatus` — MFA status + method priority for licensed users.
- `Get-TenantCalendarAccess` — delegated calendar perms, classified by mailbox type.
- `Get-TenantMailboxStorage` — usage vs quota + archive on/off and consumption.

### Administration modules (`Modules/Administration/`)
- `Add-/Remove-DistroMember` — bulk DL membership.
- `Add-BulkContact` — bulk external mail contacts.
- `Add-BulkGuestUser` — bulk B2B guest invites (auto-detect redirect URL).
- `Add-/Remove-CalendarPermission` — calendar delegate perms.
- `Add-/Remove-MailboxAccess` — FullAccess + SendAs (Add creates shared if missing).
- `Remove-UserAccess` — audit + selective removal of access/memberships.
- `Request-OneDriveProvision` — bulk OneDrive pre-provisioning (SPO).
- `Invoke-UserOffboarding` — full offboarding pipeline (see Architecture).
- `Invoke-RBACBuilder` — EXO RBAC-for-Applications scoping; group-centric,
  idempotent, template-aligned. Scale access by adding mailboxes to the group.
- `Remove-RBACBuilder` — full teardown of an app's RBAC scoping (role
  assignments, scope, EXO service principal, access group); type-to-confirm.
- `New-SelfSignedCertificate` — .pfx/.cer generator for app cert-based auth
  (CSP provider, required for EXO app-only).

All Administration/Utility/Reporting modules above are registered in the
launcher catalog and appear in the GUI.

### Utility modules (`Modules/Utility/`)
- `Install-ExchangeOnlineModule`, `Install-MicrosoftGraphModule`,
  `Install-PowerShell7`, `Install-All`.

---

## In Progress / Open items

### Invoke-UserOffboarding hardening (started 2026-09-30, from Invoke-UserOffboarding-FixPlan.md)
Scope: repo module only (`Modules/Administration/Invoke-UserOffboarding.ps1`).
Log -> `C:\Logging\MSP-M365-Utility\`; CSV stays in `C:\MSP-M365-Utility\`
(launcher View Results unaffected).
Status: code complete 2026-09-30, parses clean, ASCII-only. NOT yet run
against a live tenant - needs the section 5 testing checklist before it is
trusted in production.
- [x] 2.1 CRITICAL - licenses now removed only when safe: no mailbox (safe),
  or confirmed SharedMailbox (verified via RecipientTypeDetails) with a known
  size <= cap. Conversion-failed/unverified, "size unknown", and >50 GB all
  SKIP removal and flag RETAINED. "Already shared" treated as success.
- [x] 2.2 HIGH - empty catches removed: SendAs lookup, FullAccess index build,
  SKU map, and mailbox-stats failures are now recorded (WARN / Partial /
  Failed) instead of silently passing as Success.
- [x] 2.3 HIGH - `-BypassSecurityGroupManagerCheck` added to
  `Remove-DistributionGroupMember`.
- [x] 2.4 MEDIUM - shared-mailbox FullAccess index (`$FullAccessIndex`) built
  once before the loop; per-user removal is now a hashtable lookup. NOTE: the
  index keys on the stored grantee string (UPN); if a grant is stored under a
  non-UPN identity the lookup could miss it - watch for this during testing.
- [x] 2.6 LOW - summary shows Submitted / Processed / Skipped / Not Found.
- [x] `Write-Log` added (INFO/WARN/ERROR, `yyyy-MM-dd HH:mm:ss`, console+file,
  `-NoConsole` for file-only rows); log ->
  `C:\Logging\MSP-M365-Utility\COMPUTER_yyyyMMdd_HHmmss_UserOffboarding.log`;
  CSV renamed to `COMPUTER_yyyyMMdd_HHmmss_UserOffboarding_<Tenant>.csv`
  (still in `C:\MSP-M365-Utility\`).
- [x] Section dividers normalized; description comments added above Write-Log,
  Add-Result, ConvertTo-GB.
- [x] (2026-10-06 amendment) Name-logging: the 'Remove Group Memberships' and
  'Remove Shared Mailbox Access' rows now list the actual group names (tagged
  DL / Mail-Enabled Security / M365 / Security) and shared-mailbox addresses
  (SendAs + FullAccess), comma-separated, in the results CSV Detail column.
  Add-Result gained a `-KeepDetail` switch so these two rows keep their Detail
  on Success (the blank-on-Success convention still holds everywhere else).
Not in this pass: 2.5 (#Requires vs install block), 2.7 (redundant revoke),
2.8 (owned groups / forwarding default / prompt ordering).


- [x] ~~Offboarding not yet run against a live tenant.~~ **Validated** against a
  live tenant on 2026-09-04 (used by the operator, worked as intended).
- [ ] **RBAC Builder/Removal not yet run against a live tenant.** Reworked but
  unverified end-to-end; test the create → add-to-group → remove lifecycle on a
  non-production app first.
- [x] ~~`Microsoft.Graph.Applications` not in the prereq installer.~~ Done in
  0.12.1-beta: added to the MS Graph pill (now 6 submodules),
  Install-MicrosoftGraphModule, and Install-All.

---

## Known constraints / gotchas

- **MSAL conflict:** old `ExchangeOnlineManagement` (pre-3.5) bundles an old MSAL
  that breaks `Connect-MgGraph` if EXO loads first — even under PS7. Mitigated by
  connecting Graph first everywhere. Keeping EXO updated is the real fix.
- **SharePoint Online Management Shell** (used by `Request-OneDriveProvision`) is
  most reliable under Windows PowerShell 5.1.
- **Sign-in activity / Stale** (`Get-TenantMailboxes`) requires Entra ID P1/P2
  (tenant has P2).
- **Repo/commit identity:** pushes use `gh` account **MasatoNakajima20**; commits
  are authored as `Masato Nakajima <anonymousname20@gmail.com>` (intentionally
  separate from `john@3foldit.com` / the org).
