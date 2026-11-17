# Email Flow Agent Guide (PM & Re-Auth Projects)

Reusable implementation guide for adding the Pending Pickup-style email workflow to the **PM** and **Re-Auth** report projects. Adapt names, sheet names, ranges, and copy to each project; keep the same architecture and safeguards.

## 1. Goals

- Users configure email **To/CC/Subject/Body** in dedicated hidden template tabs (not buried in code).
- **Manager/recipient-specific** emails paste the correct tab/range; **director/summary** email pastes the dashboard/summary range.
- Drafts open in **Outlook** for review/send (no auto-send), preserving Excel formatting via paste into Word editor.
- Macro-enabled test template and no-macro/work-PC artifacts stay aligned.
- Setup is resilient: legacy fallbacks, value migration, clear errors/logging.

## 2. Architecture Overview

| Layer | Responsibility |
|---|---|
| **Hidden config sheets** (template) | User-editable TO/CC/SUBJECT/BODY; formatted UI, instructions, Return button |
| **`modDeclarations.bas`** | Constants: sheet names, config cell addresses, default subjects/bodies, legacy fallback cells |
| **`modDialog.bas`** (or email module) | Read config, setup/return subs, range capture, Outlook compose/paste, button handlers |
| **Report builder (`modCore`/formatting)** | Create per-recipient and dashboard email buttons on generated report; set `AlternativeText` to target sheet name |
| **`scripts/sync_template.py`** | Sync VBA into `.xlsm`, provision/format config sheets, migrate/preserve values, rebuild no-macro `.xlsx` + build ZIP |
| **Run/watch tooling** | Optional: detect config changes and sync before runs / on save so artifacts stay matched |

## 3. Email Template Sheets (UI/Config)

Create one hidden sheet per email context (examples: Manager, Director/Dashboard; PM/Re-Auth may differ).

**Recommended layout (current standard):**
- `D5` = **TO**
- `D7` = **CC** (spacer row above for separation)
- `D9` = **SUBJECT**
- `D11` = **BODY** (merged range, e.g. `D11:K31`)
- Instructions row (e.g. `B4`): `TO/CC: separate addresses with semicolons (;)  •  BODY: keep {{REPORT}} for pasted report (text above = intro, below = closing); use Alt+Enter for a new line.`
- **Return to START HERE** button in header: unhides START HERE, activates it, hides the config sheet.
- Shape-free cell formatting for inputs (borders/fills/fonts) so typed text is never obscured by shapes. Hide gridlines/headings/unused rows/cols for a clean view.

**Preserve cell addresses once chosen** (VBA + sync + watchers depend on them). If changing layout, add migration logic.

## 4. `modDeclarations.bas` Constants

Define per template:
```vb
Public Const EMAIL_CONFIG_SHEET As String = "Manager Email Template" ' adjust per project
Public Const EMAIL_CONFIG_TO_CELL As String = "D5"
Public Const EMAIL_CONFIG_CC_CELL As String = "D7"
Public Const EMAIL_CONFIG_SUBJECT_CELL As String = "D9"
Public Const EMAIL_CONFIG_BODY_CELL As String = "D11"
' Director/dashboard equivalents...
' Legacy START HERE fallback cells if migrating from inline config
Public Const EMAIL_SUBJECT_MANAGER As String = "..."  ' project-specific default
Public Const EMAIL_SUBJECT_DIRECTOR As String = "..." ' project-specific default
```
Keep names project-specific but consistent; document them.

## 5. VBA: Read, Setup, Return

- **Readers:** `GetEmailTo/CC/Subject/Body` and director variants.
  - Read from hidden template sheet cells; trim; on error/blank subject/body return sensible defaults.
  - Prefer template sheet; optionally fall back to legacy START HERE cells for backward compatibility.
  - Log reads (abbreviate body) for diagnostics.
- **Setup subs** (wired to START HERE buttons): `SetupEmailCC` / `SetupDirectorEmail` (rename per project).
  - Unhide template sheet, activate, select TO cell.
  - Show concise `MsgBox` instructions: which cells, semicolon for multiple addresses, blank = default where applicable, and how to return/re-hide.
- **Return:** `Public Sub ReturnToStartHere()` (or per-sheet safe variant).
  - Resolve caller sheet (Application.Caller shape parent) with fallback if active sheet is a known email template.
  - Show/activate START HERE, hide template sheet. Log action.

## 6. VBA: Compose & Send (Outlook)

**Core pattern (reliable):**
1. Determine target worksheet/range.
2. `range.Copy` (preserve formatting).
3. `CreateObject("Outlook.Application")` -> `CreateItem(0)`, set To/CC/Subject from readers.
4. `.Display` + `DoEvents`, wait/retry for `GetInspector.WordEditor` (e.g. up to ~8s) — editor can be slow.
5. Build body in Word range:
   - Greeting/intro text
   - Paste copied Excel range at insertion point
   - Closing/signature text
6. Leave draft displayed for user review/send. Clear `Application.CutCopyMode` in cleanup. User-facing `MsgBox` + logging on failures.

**Range capture rules:**
- **Manager/recipient tab:** copy the data region (e.g. header row to last used row/columns) — exclude UI buttons placed outside data columns.
- **Dashboard/summary:** compute bounds including `UsedRange` and embedded `ChartObjects`/`Shapes` so charts are included.

**Body/`{{REPORT}}` marker (recommended UX):**
- Treat `{{REPORT}}` (or `{{TABLE}}`) in BODY as the paste location: text above = intro, below = closing.
- If marker missing, define explicit fallback (e.g. treat body as intro, paste after, append default closing) and document it.
- In Excel template cell, new lines via **Alt+Enter**; preserve line breaks when composing.
- *Implementation note:* if not yet implemented, keep behavior consistent and document; adding a `ComposeEmailBody` helper is a clean upgrade.

**Subject/date rules:** apply project policy (e.g. append current date for monthly director/dashboard subject if not present; leave custom subjects as-is when appropriate).

## 7. Button Wiring on Generated Reports

- **Per-manager/recipient button** on each generated tab:
  - `OnAction` -> handler (e.g. `EmailSelectedManagerTab`)
  - `AlternativeText` = target sheet name (stable identifier; do not rely on mutable globals overwritten in loops)
  - Handler: resolve `Application.Caller` shape -> `AlternativeText` -> worksheet in report workbook -> compute row/context -> `SendManagerEmail`
- **Dashboard/director button:** `OnAction` -> `EmailDashboard`, `AlternativeText` = dashboard sheet name (with fallback lookup).
- Place buttons outside copied data ranges; size/position consistently.

## 8. `sync_template.py` Responsibilities

- Import all project VBA modules into `.xlsm` (exclude legacy monolith if applicable); ensure START HERE/setup buttons exist and point to correct macros.
- **Provision + format** each email config sheet:
  - Create if missing, set hidden, clear/rebuild clean layout (remove stray shapes where inputs must be unobstructed).
  - Read existing values with **migration** across old cell layouts (detect by label cells/known positions) and preserve user-entered TO/CC/SUBJECT/BODY.
  - Normalize only known obsolete defaults if required; avoid clobbering custom content.
  - Write values back to canonical cells and apply formatting/instructions.
- Save `.xlsm`, regenerate **no-macro `.xlsx`** and **`Project-Local-Build.zip`** (unless explicitly skipped) so work-PC artifacts match.

## 9. Keeping Macro and No-Macro Artifacts Aligned

- **Source of truth for VBA:** `.bas` modules + sync.
- **Source of truth for user email config (typical):** `.xlsm` template tabs; sync propagates to regenerated no-macro/ZIP.
- Provide automation where possible:
  - `run-report.py` (or project equivalent): compare key config cells (`D5/D7/D9/D11` per sheet) between `.xlsm` and no-macro before run; run sync if different/unreadable.
  - Optional watcher: detect stable saves to `.xlsm`, compare same cells, auto-sync.
- Document workflow: edit/test in macro template -> compile/sync -> distribute updated no-macro/ZIP.

## 10. Validation Checklist (per project)

- [ ] `test-compile.py` (or equivalent) passes after VBA changes.
- [ ] `sync_template.py` provisions sheets, preserves custom values, migrates legacy layouts, rebuilds artifacts.
- [ ] Setup buttons open correct hidden sheet/cell; **Return to START HERE** re-hides correctly.
- [ ] Manager/recipient email pastes correct tab range; dashboard email pastes dashboard incl. charts.
- [ ] To/CC/Subject/Body read from template cells; defaults/date rules behave as specified.
- [ ] Outlook draft opens with formatting intact; missing-Outlook/editor-timeout errors are user-friendly and logged.
- [ ] `.xlsm`, no-macro `.xlsx`, and build ZIP aligned after edits.
- [ ] README/BUILD-INSTRUCTIONS updated (cells, `{{REPORT}}`/Alt+Enter, semicolons, workflow).

## 11. Adaptation Notes for PM & Re-Auth

- Rename sheets/cells/constants/macros/buttons to match each domain (e.g. PM manager summaries, re-auth queues) but keep the same cell contract and handler pattern.
- Decide recipients/contexts (one vs multiple template tabs) and defaults/subject date policy per report.
- Ensure snapshot ranges match each report’s layout (headers, KPI bands, charts).
- If reports are generated workbooks, confirm buttons resolve sheets in the **output workbook** (not the template) via `Application.Caller`/`AlternativeText`.
- Start by porting: declarations + readers/setup/return + compose/send + sync provisioning + button wiring, then tailor formatting/copy.

## 12. Key Pitfalls to Avoid

- Storing target sheet in a global overwritten during generation — use `AlternativeText` per button.
- Shapes overlaying editable cells — prefer cell-formatted inputs.
- Not waiting for Outlook Word editor — causes paste/compose failures.
- Copy ranges that include action buttons or miss charts — compute bounds intentionally.
- Sync that overwrites user config — always migrate/preserve values.
- Forgetting to regenerate/ship no-macro + ZIP after template/config changes.
