# ARCHITECTURE.md — MSP M365 Utility

How the system is put together. Update when structure, components, or data flow change.

## Overview

MSP M365 Utility is a **single-file WinForms launcher** plus a set of standalone
PowerShell **module scripts** that each perform one Microsoft 365 report or
administrative action. It is designed to run on any Windows machine with no
local clone:

```
iex (irm "https://raw.githubusercontent.com/MasatoNakajima20/MSP-M365-Utility/main/Launch-MSPM365Utility.ps1")
```

There is no compiled build and no shared runtime library — each module is
self-contained and self-installs its own prerequisites. The launcher is a
dispatcher/UI, not a framework the modules depend on.

## Components

### 1. Launcher — `Launch-MSPM365Utility.ps1` (repo root)

- WinForms GUI. Holds `$script:Version` (the released version shown in the title
  bar and About dialog).
- Holds `$script:Modules` — the **catalog**: one object per module with
  `File` (repo-relative path), `Title`, `Category`, `Description`. This catalog,
  not the folder contents, is what the GUI renders. **A module file on disk that
  is not in this catalog does not appear in the GUI.**
- Landing screen = category tiles (**Reporting / Administration / Utility**) +
  a prerequisite status banner. Drill into a category to see its module cards;
  each card's **Run** button invokes the module.
- `Get-RequiredPrereqStatus` / `Get-LatestModuleVersion` — compute the pill
  states by scanning module versions across **all** standard module roots on
  disk (PS7 + Windows PowerShell, CurrentUser + AllUsers), so a module updated
  under one shell scope is still detected.
- `Invoke-Module` — downloads the selected module from the GitHub **raw** URL
  into `%TEMP%\MSPM365Utility\` and launches it with `pwsh` (falling back to
  `powershell`) in a **new window**, no `-NoExit`.

### 2. Modules — `Modules/<Category>/*.ps1`

Three category folders, mapped 1:1 to the launcher's categories:

- **`Modules/Reporting/`** — read-only tenant exports (CSV).
- **`Modules/Administration/`** — actions that change tenant state.
- **`Modules/Utility/`** — local prerequisite installers.

Every module follows the same **template**: banner → tenant-code prompt →
optional inputs (paste-list / scope) → parse/validate → (for admin) preview +
confirmation → module check/connect → progress-bar processing loop → CSV export
→ disconnect → summary → "Press Enter to exit".

### 3. Output — `C:\MSP-M365-Utility\`

All report and results CSVs are written here, named
`<Report>_<TENANTCODE>_<yyyyMMdd_HHmmss>.csv`. The launcher's **View Results**
button opens this folder. (The module *cache* is separate: `%TEMP%\MSPM365Utility\`.)

## Data flow

```
 Operator ──run one-liner──▶ Launcher GUI
     │                          │
     │  clicks Run on a card    │ downloads Modules/<cat>/<name>.ps1
     ▼                          ▼  from GitHub raw ──▶ %TEMP%\MSPM365Utility\
 New PowerShell window ◀── launched with pwsh/powershell
     │
     │  prompts (tenant code, paste-list, confirmations)
     ▼
 Connect: Microsoft Graph FIRST, then Exchange Online / SharePoint Online
     │
     ▼
 Process ──▶ CSV to C:\MSP-M365-Utility\  +  on-screen summary
```

## Service dependencies

- **Microsoft Graph** (`Microsoft.Graph.*` submodules): Users, Users.Actions,
  Groups, Reports, Identity.SignIns. Used by user/group/MFA/offboarding modules.
- **Exchange Online** (`ExchangeOnlineManagement`): mailbox, calendar, DL,
  permission, and storage modules.
- **SharePoint Online** (`Microsoft.Online.SharePoint.PowerShell`):
  `Request-OneDriveProvision` only.

### Connection-order rule (important design decision)

Modules that use **both Graph and EXO connect to Graph first**. EXO bundles an
older `Microsoft.Identity.Client` (MSAL); if it loads first, `Connect-MgGraph`
fails with `Method not found: WithLogging` — even under PowerShell 7, whose
assembly isolation does not cover this auth path. Loading Graph's newer MSAL
first avoids the clash. Graph modules are also listed first in each module's
import loop.

## Key design decisions

- **Self-contained modules, catalog-driven GUI.** Adding a module = drop a
  `.ps1` in the right category folder **and** add an entry to `$script:Modules`
  in the launcher. The folder and the catalog are maintained together.
- **Fetch-on-demand over bundling.** The launcher pulls each module live from
  `main`, so a push updates all users instantly with no redistribution.
- **Standardized output path** (`C:\MSP-M365-Utility\`) so reports are always
  found in one place regardless of which module produced them.
- **Confirmation depth scales with blast radius.** Read-only reports just run;
  bulk admin modules preview + confirm; `Invoke-UserOffboarding` gates twice
  (type-the-count, then per-account Y/N).
- **Exceptions-style logging** (offboarding): every action is logged, but
  Success rows carry a blank detail; the detail is reserved for
  Partial/Info/Skipped/Failed.

## Notable module: `Invoke-UserOffboarding`

Ordered per-account pipeline (Graph + EXO): disable account → remove manager →
revoke sessions → wipe MFA methods → revoke tokens → convert to shared → optional
forwarding → optional delegate grants (FullAccess + SendAs) → remove group
memberships + shared-mailbox access → mailbox size gate (retain license if
> 50 GB, else remove all licenses one SKU at a time). Emits a per-user/per-action
results CSV.

## Notable capability: RBAC-for-Applications lifecycle

A three-module lifecycle for scoping a Graph app's *application* permissions to a
subset of mailboxes (the go-forward alternative to the deprecated Application
Access Policy):

- **`Invoke-RBACBuilder`** creates an EXO service principal for the app, a
  mail-enabled **security group** as the access boundary, a **management scope**
  filtered on that group's membership, and **role assignment(s)** derived by
  mapping the app's consented Graph permissions (Mail.Read, Mail.Send, …) to EXO
  "Application X" roles. It is idempotent (reuses existing objects) and seeds the
  group with the pasted mailboxes.
- **Scaling** is by group membership: to grant another mailbox, add it to
  `<App>_RBAC_Access` (e.g. via `Add-DistroMember`) — no re-run of the builder.
- **`Remove-RBACBuilder`** discovers the objects by naming convention +
  role-assignee match and tears them down in dependency order (role assignments →
  scope → service principal → group), behind a type-the-app-name confirmation.

`New-SelfSignedCertificate` complements this by producing the `.pfx`/`.cer`
keypair (CSP provider) an app uses for certificate-based auth.

## Repository conventions

- **Categories are folders.** Keep `Modules/<Category>/` as the source of truth
  for placement; do not reorganize without instruction.
- **Secrets never committed** — `.gitignore` excludes `.pfx/.cer/.key/.env/secret*`;
  the repo is public, so scripts prompt for secrets at runtime rather than
  embedding them.
- **ASCII-only** in scripts (no Unicode box-drawing / em dashes); banners use
  `+`, `-`, `|`.
