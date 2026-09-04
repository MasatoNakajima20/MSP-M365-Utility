# PROJECT_STATE.md — MSP M365 Utility

Living status log for the MSP M365 Utility. Update after every meaningful change.

- **Repo:** https://github.com/MasatoNakajima20/MSP-M365-Utility (public)
- **Local path:** `C:\Claude Projects\MSP 365 Reporting Tool`
- **Current tagged release:** `0.12.1-beta`
- **Launcher `$script:Version`:** `0.12.1-beta`
- **Last updated:** 2026-09-04

---

## Versioning rule (standing)

Applied on every push:

- **Patch (0.0.x)** — a fix or small adjustment.
- **Minor (0.x.0)** — a new module, or a change set over 500 lines.
- **Major (x.0.0)** — not yet defined; ask before bumping.
- If a push contains both a fix and a new module, the **minor** bump wins.
- After each push: compute the next version, set `$script:Version` in the
  launcher, then `git tag -a <ver> -m <ver>; git push origin <ver>`.
- Beta suffix stays until told otherwise.

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
  City/State/Country/phones and AccountStatus.
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
