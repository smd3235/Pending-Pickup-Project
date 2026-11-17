# Pending Pickup Email Implementation — Review & Correct Implementation Guide

Project-specific review of the email feature as implemented in this repository. This document reflects the current Pending Pickup architecture (hidden Manager/Director Email Template tabs, START HERE setup buttons, per-tab report buttons, Outlook Word-editor paste) and should be treated as the source for email changes.

**Do not copy** the PM Requests/hidden-sheet email pattern (`EmailTemplate`/`EmailSettings`, HTML template placeholder, `GetEmailTemplate`, image Content-ID attachments) into this project unless explicitly redesigned; `documentation/EMAIL_FEATURE_PORTING_GUIDE.md` describes a different project pattern.

---

## 1. Current architecture

### Configuration

#### Manager emails (hidden template tab, `Manager Email Template`)
Defined in `modDeclarations.bas` and read by `modDialog.bas`:

| Constant | Value | Purpose |
|---|---|---|
| `EMAIL_CONFIG_SHEET` | `Manager Email Template` | Hidden manager email config sheet |
| `EMAIL_CONFIG_TO_CELL` | `D5` | To recipients (comma-separated) |
| `EMAIL_CONFIG_CC_CELL` | `D6` | CC recipients (comma-separated) |
| `EMAIL_CONFIG_SUBJECT_CELL` | `D7` | Subject override (blank = default) |
| `EMAIL_CONFIG_BODY_CELL` | `D9` | Intro/body text override (blank = default) |

- `GetEmailTo()` reads `D5`; `GetEmailCC/Subject/Body()` read `D6/D7/D9` with **legacy fallback** to START HERE `D16-D18` for older templates.
- `SetupEmailCC()` unhides/activates `Manager Email Template` and selects `D5` (TO) for editing.
- Template layout (provisioned by `scripts/sync_template.py`): labeled sections **TO/CC/SUBJECT/BODY**, clean gridline-free view with only input areas visible, and a large merged body input (`D9:K24`).

#### Director/dashboard email (hidden template tab, `Director Email Template`)
Defined in `modDeclarations.bas` and read by `modDialog.bas`:

| Constant | Value | Purpose |
|---|---|---|
| `DIRECTOR_EMAIL_SHEET` | `Director Email Template` | Hidden director/dashboard email config sheet |
| `DIRECTOR_EMAIL_TO_CELL` | `D5` | To recipients (comma-separated) |
| `DIRECTOR_EMAIL_CC_CELL` | `D6` | CC recipients (comma-separated) |
| `DIRECTOR_EMAIL_SUBJECT_CELL` | `D7` | Subject override (blank = default) |
| `DIRECTOR_EMAIL_BODY_CELL` | `D9` | Intro/body text override (blank = default) |

- `GetDirectorEmailTo/CC/Subject/Body()` read from `ThisWorkbook` (template), with trim/default fallback behavior.
- `SetupDirectorEmail()` unhides/activates the director template sheet and selects `D5` (TO) for editing.
- Same clean visual layout as manager template: **TO/CC/SUBJECT/BODY** with large merged body input (`D9:K24`) and hidden non-input areas/gridlines.

`scripts/sync_template.py` ensures both hidden sheets, labels, and buttons exist in the template/build artifacts, and START HERE shows pointers to those template tabs.

### Button creation

- **START HERE (template, `scripts/sync_template.py`)**: `btnSetupEmailCC` → `SetupEmailCC` (“Setup Manager Email Settings”), `btnSetupDirectorEmail` → `SetupDirectorEmail` (“Setup Director Email Template”).
- **Manager tabs (`modCore.bas`)**: shape per manager, `Text` `"Email " & manager`, `OnAction` `"EmailSelectedManagerTab"`, `AlternativeText` = target sheet name.
- **Dashboard (`modCore.bas`)**: `btnEmailDashboard`, `OnAction` `"EmailDashboard"`, `AlternativeText` = dashboard sheet name.

### Click handlers (`modDialog.bas`)

- `EmailSelectedManagerTab()`: resolves caller via `Application.Caller`/`ResolveCallerShape`, resolves target worksheet from the owning workbook (`shp.Parent.Parent`) using `AlternativeText`, derives manager name from button text (fallback sheet name), computes `rowCount = LastUsedRow - MANAGER_HEADER_ROW`, and calls `SendManagerEmail`.
- `EmailDashboard()`: resolves target from caller/fallback `Dashboard`, estimates row count from the manager breakdown table when available, and calls `SendDashboardEmail`.

### Send implementation (`modDialog.bas`)

- **Delivery method:** copy Excel range to clipboard, open Outlook inspector (`.Display`), paste into `mail.GetInspector.WordEditor` to preserve formatting. No temp image files.
- **Outlook:** late-bound `CreateObject("Outlook.Application")`, `CreateItem(0)`, `GetWordEditor(mail, 8)` retry loop, `DoEvents` while waiting.
- **Manager range:** `GetManagerSnapshotRange` captures `A{MANAGER_HEADER_ROW}` through last data row and header-bounded columns to avoid copying button columns.
- **Dashboard range:** `GetDashboardSnapshotRange` computes bounds from used cells and embedded chart/shape extents (not a fixed window).
- **Manager draft:** `To`/`CC`/Subject/Intro from `Manager Email Template` (`D5/D6/D7/D9`) with legacy START HERE fallback and default fallbacks; pasted range + standard closing.
- **Director draft:** `To`/`CC`/Subject/Intro from `Director Email Template` (`D5/D6/D7/D9`) with default fallbacks (default subject includes report date), pasted dashboard range, standard closing.
- **Cleanup/UX:** `Application.CutCopyMode = False` cleared on exit; copy/editor/Outlook failures are logged and surfaced with `MsgBox` via `NotifyEmailFailure`.

---

## 2. Remaining gaps / follow-ups

1. **Subject cells are full overrides:** manager `D6` and director `D7` replace generated defaults (no placeholder tokens). Document if templating is later desired.
2. **Config is template-scoped:** email settings are read from `ThisWorkbook`, while drafts are created for sheets in the clicked report workbook; correct when macro/template context owns config, but worth noting if distribution model changes.
3. **Template tab visibility:** setup routines unhide tabs for editing; users may leave them visible. Sync sets them hidden by default.
4. **Dashboard row-count heuristic:** uses manager breakdown rows when found, else falls back to sheet used rows; refine if a more explicit pending-count source is added.
5. **No end-to-end automation for Outlook UI:** unattended runs cannot exercise inspector/Word-editor paste; manual verification remains required.

---

## 3. Manual verification scenarios

1. **Manager tab:** click `Email <Manager>` → Outlook draft opens with To/CC/Subject/Intro from `Manager Email Template` (`D5/D6/D7/D9`) or defaults/legacy fallback, correctly bounded table paste, closing present.
2. **Dashboard:** click dashboard email → draft opens with To/CC/Subject/Intro from `Director Email Template` (`D5/D6/D7/D9`) or defaults, full dashboard snapshot including KPI/table/chart area, correct closing.
3. **Config overrides:** set/clear manager template `D5/D6/D7/D9` and director template `D5/D6/D7/D9`; confirm defaults/fallbacks.
4. **Setup buttons:** START HERE “Setup Manager Email Settings” and “Setup Director Email Template” open/unhide the correct template tabs (TO selected) showing clean TO/CC/SUBJECT/BODY layout with large body area.
5. **Failure modes:** missing Outlook or delayed editor → logged and user-visible `MsgBox`, no hang/crash, clipboard cleared.

---

## 4. Related files

- `modDialog.bas` — config readers, setup routines, range helpers, send routines, button handlers
- `modCore.bas` — report-sheet email button creation and `AlternativeText` targets
- `modDeclarations.bas` — `EMAIL_*`, legacy manager fallback, and `DIRECTOR_EMAIL_*` constants
- `modLogging.bas` — `LogMsg` / `LogErr`
- `scripts/sync_template.py` — template button/config/sheet provisioning and build package rebuild
- `Pending Pickup Template.xlsm` / `Template (No Macros).xlsx` / `Project-Local-Build.zip` — template/build artifacts
- `scripts/test-compile.py`, `scripts/run-report.py` — validation
- `documentation/EMAIL_FEATURE_PORTING_GUIDE.md` — different (PM Requests) pattern; not authoritative for this project
- `documentation/PORTABILITY_AND_WORK_PC_PATHS.md` — separate portability issue, not email-specific
