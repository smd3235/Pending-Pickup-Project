# Pending Pickup Email Implementation — Review & Correct Implementation Guide

Project-specific review of the email feature as implemented in this repository. This document reflects the current Pending Pickup architecture (START HERE config cells, per-tab buttons, Outlook Word-editor paste) and should be treated as the source for email changes.

**Do not copy** the PM Requests/hidden-sheet email pattern (`EmailTemplate`/`EmailSettings`, HTML template placeholder, `GetEmailTemplate`, image Content-ID attachments) into this project unless explicitly redesigned; `documentation/EMAIL_FEATURE_PORTING_GUIDE.md` describes a different project pattern.

---

## 1. Current architecture

### Configuration (template-driven, `START HERE`)
Defined in `modDeclarations.bas` and read by `modDialog.bas`:

| Constant | Value | Purpose |
|---|---|---|
| `EMAIL_CONFIG_SHEET` | `START HERE` | Sheet containing email settings |
| `EMAIL_CONFIG_CC_CELL` | `D16` | CC recipients (comma-separated) |
| `EMAIL_CONFIG_SUBJECT_CELL` | `D17` | Subject override (blank = default) |
| `EMAIL_CONFIG_BODY_CELL` | `D18` | Intro message override (blank = default) |

- `GetEmailCC()` reads `D16` (trimmed) from `ThisWorkbook`.
- `GetEmailSubject(defaultSubject)` reads `D17`; falls back to `defaultSubject` when blank/error.
- `GetEmailBody(defaultBody)` reads `D18`; falls back to `defaultBody` when blank/error and is used as the draft intro text.
- `SetupEmailCC()` activates `START HERE!D16` and explains `D16/D17/D18` editing.

### Button creation (`modCore.bas`, during report build)
- **Manager tabs:** rounded-rectangle shape placed at `managerLastCol + 1`, row `MANAGER_HEADER_ROW`.
  - `Text`: `"Email " & manager`, `OnAction`: `"EmailSelectedManagerTab"`, `AlternativeText`: target sheet name.
- **Dashboard:** shape near column `Q`, row 2.
  - `Text`: `Email Dashboard to Director`, `OnAction`: `EmailDashboard`, `AlternativeText`: dashboard sheet name.

### Click handlers (`modDialog.bas`)
- `EmailSelectedManagerTab()`: resolves caller via `Application.Caller`/`ResolveCallerShape`, resolves the target worksheet from the owning workbook (`shp.Parent.Parent`) using `AlternativeText`, derives manager name from button text (fallback sheet name), computes `rowCount = LastUsedRow - MANAGER_HEADER_ROW`, and calls `SendManagerEmail`.
- `EmailDashboard()`: resolves target from caller/fallback `Dashboard`, estimates row count from the manager breakdown table when available, and calls `SendDashboardEmail`.

### Send implementation (`modDialog.bas`)
- **Delivery method:** copy Excel range to clipboard, open Outlook inspector (`.Display`), paste into `mail.GetInspector.WordEditor` to preserve formatting. No temp image files.
- **Outlook:** late-bound `CreateObject("Outlook.Application")`, `CreateItem(0)`, `GetWordEditor(mail, 8)` retry loop, `DoEvents` while waiting.
- **Manager range:** `GetManagerSnapshotRange` captures `A{MANAGER_HEADER_ROW}` through last data row and header-bounded columns to avoid copying button columns.
- **Dashboard range:** `GetDashboardSnapshotRange` computes bounds from used cells and embedded chart/shape extents (not a fixed window).
- **Subjects:** manager default includes manager name and pending count; dashboard default includes report date. `D17` overrides defaults.
- **Body:** `"Hi <name>,"` + configured/default intro (`D18`) + pasted range + closing `"Please review and respond as needed."` / `"Pending Pickup Report System"`.
- **Cleanup/UX:** `Application.CutCopyMode = False` cleared on exit; copy/editor/Outlook failures are logged and surfaced with `MsgBox` via `NotifyEmailFailure`.

---

## 2. Remaining gaps / follow-ups

1. **`D17` is a full override:** subject cell replaces the generated default (no placeholder tokens). Document if templating is later desired.
2. **Config is template-scoped:** email settings are read from `ThisWorkbook` (`START HERE`), while drafts are created for sheets in the clicked report workbook; this is correct when macro/template context owns config, but worth noting if distribution model changes.
3. **Dashboard row-count heuristic:** uses manager breakdown rows when found, else falls back to sheet used rows; refine if a more explicit pending-count source is added.
4. **No end-to-end automation for Outlook UI:** unattended runs cannot exercise inspector/Word-editor paste; manual verification remains required.
5. **Range edge cases:** heavily customized layouts or unusual floating objects may still need bounds tuning; validate against real report output.

---

## 3. Manual verification scenarios

1. **Manager tab:** click `Email <Manager>` → Outlook draft opens, To empty, CC from `D16`, subject includes manager and pending count (or `D17` override), intro from `D18`/default, correctly bounded table paste, closing present.
2. **Dashboard:** click dashboard email → draft opens with full dashboard snapshot including KPI/table/chart area, correct CC/subject/body.
3. **Config overrides:** set/clear `D16/D17/D18` on `START HERE` and confirm defaults/fallbacks.
4. **Multiple managers:** different buttons open correct sheets/ranges/names (no stale target).
5. **Failure modes:** missing Outlook or delayed editor → logged and user-visible `MsgBox`, no hang/crash, clipboard cleared.

---

## 4. Related files

- `modDialog.bas` — config readers, setup, range helpers, send routines, button handlers
- `modCore.bas` — email button creation and `AlternativeText` targets
- `modDeclarations.bas` — `EMAIL_*` constants/config cell addresses
- `modLogging.bas` — `LogMsg` / `LogErr`
- `Pending Pickup Template.xlsm` / `Template (No Macros).xlsx` / `Project-Local-Build.zip` — template/build artifacts
- `scripts/sync_template.py`, `scripts/test-compile.py`, `scripts/run-report.py` — build/validation
- `documentation/EMAIL_FEATURE_PORTING_GUIDE.md` — different (PM Requests) pattern; not authoritative for this project
- `documentation/PORTABILITY_AND_WORK_PC_PATHS.md` — separate portability issue (hardcoded/dev-machine paths), not email-specific
