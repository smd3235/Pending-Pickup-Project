# Changelog

All notable changes to the Pending Pickup Report Builder. Each entry is stamped
`YYYY-MM-DD HH:MM:SS` (local time). Version `1.0` is the current baseline.

## 2026-11-18

- **15:25:00** — Left-aligned START HERE "GOOD TO KNOW" section header (B24:E24) and all four note bullets (B25:E25 through B28:E28): set `HorizontalAlignment = -4131` (left) and `IndentLevel = 0` on each, instead of center/indent 1. Re-synced `Pending Pickup Template.xlsm`, both ZIPs rebuilt.

- **15:20:00** — Fixed START HERE step alignment: steps 1-5 now merge B:E per row (was D:E) so they sit directly under the "STEP 3 • HOW IT WORKS" header (B17:E17), no longer shifted to column D. Re-synced `Pending Pickup Template.xlsm`, both ZIPs rebuilt.

- **15:05:00** — Fixed three START HERE page issues in `scripts/sync_template.py`: (1) Button labels were missing because `_make_button` only set text via `TextFrame2` with no `Font.Color` fallback — added `Font.Color = 0xFFFFFF` and a regular `TextFrame` fallback when `TextFrame2` fails. (2) Step 3 text (1-5) was drifting off to the right because numbers lived in column D and text in column E as separate cells — now merges D:E per step row with `num + "  " + text` so they stay aligned directly below each other. (3) Start macro button was auto-running without file picker — deleted `log/auto-source.txt` and `log/auto-skip-prior.txt` so clicking CREATE in the template prompts for source and prior files (automated `run-report.py` still recreates these files for unattended runs).

- **13:42:00** — Fixed `Dir$` infinite loop in `vba/modCore.bas` VBA import block: `Dir$` was matching non-`.bas` entries (e.g. `.gitkeep` via Windows 8.3 short-name mapping) and directory entries (`.`/`..`), causing repeated import attempts that hung the report run past the 30s cap and produced RPC failure. Now filters strictly by `.bas` extension via `LCase$(Right$(basFile, 4)) = ".bas"` before importing into `outWb.VBProject.VBComponents`. Report now generates macro-enabled `.xlsm` output successfully (`Reports/Pending Pickup 2026-11-18_134311.xlsm`) with completion marker written; `test-compile.py` → COMPILE OK; `sync_template.py` synced 7 modules + button; both ZIPs rebuilt.

- **13:35:00** — Fixed broken START HERE page layout in `scripts/sync_template.py`: `_make_button` helper was defined after its call sites (all inside `try/except`, so button creation silently failed and no buttons appeared). Moved `_make_button` definition before all button creation calls, rewrote layout with proper cell-anchored cards (hero band, CTA card with large CREATE button, email setup card with side-by-side manager/director buttons, numbered how-it-works steps, good-to-know panel), consistent spacing and typography. Regenerated `template-start-here.png` screenshot.

- **13:30:00** — Regenerated all QA screenshots after START HERE redesign and manager-tab unfreeze: `template-start-here.png` now reflects the new hero/CTA/email-layout design, and all 13 report screenshots (Dashboard, 10 manager tabs, Raw Data, Audit Log) were regenerated via `scripts/snapshot-report.py` to match current output.
- **13:20:00** — Unfroze manager tab panes in `vba/modFormatting.bas` (`FormatManagerSheet`): changed `ActiveWindow.SplitRow` from `MANAGER_HEADER_ROW` (16) to `0` and set `ActiveWindow.FreezePanes = False` so the entire manager tab scrolls together instead of pinning the KPI cards at the top. KPI cards still render at the top of each tab but no longer block scrolling of the data table below.

- **13:15:00** — Redesigned START HERE page layout in `scripts/sync_template.py`: new hero band (deep blue, centered title/subtitle), prominent CTA card with large "CREATE PENDING PICKUP REPORT" button, email setup card with side-by-side manager/director setup buttons, numbered "HOW IT WORKS" steps, and a clean "GOOD TO KNOW" notes panel. Replaced cramped multi-column layout with single-content-area (B:E) cards, improved spacing, consistent typography hierarchy, and a modern blue/white color palette.

- **13:05:00** — Manager email attachment privacy: `SendManagerEmail` in `vba/modDialog.bas` now creates a temp workbook containing ONLY that manager's tab when "Attach Report" (`D36` / registry `AttachReport=1`) is enabled, instead of attaching the full report with all managers' data. Added `CreateManagerOnlyWorkbook` helper that copies the manager sheet into a new macro-free `.xlsx`, saves to `%TEMP%`, attaches it, and deletes the temp file after the email is created. Prevents other managers' data from being exposed to individual managers.
- **12:50:00** — Strengthened email system in `vba/modDialog.bas` and `vba/modDeclarations.bas`: implemented `{{REPORT}}` marker body splitting (intro above marker, closing below, report pasted at marker position), added registry-driven auto-send toggle (`EMAIL_REG_AUTO_SEND`) and report attachment toggle (`EMAIL_REG_ATTACH_REPORT`), added TO address validation before opening Outlook, added `SendTestEmail` sub to verify config, and preserved Outlook signature by inserting [intro][report][closing][signature]. Both `SendManagerEmail` and `SendDashboardEmail` now accept an optional attachment path and respect auto-send vs draft mode.
- **12:50:01** — `scripts/sync_template.py` now provisions Auto Send (D35) and Attach Report (D36) config cells on both Manager and Director Email Template sheets, preserves existing values across re-syncs, and styles them as labeled toggle rows.
- **12:50:02** — `vba/modDeclarations.bas` added email constants: `EMAIL_REG_AUTO_SEND`, `EMAIL_REG_ATTACH_REPORT`, `EMAIL_MODE_DRAFT`, `EMAIL_MODE_SEND`, `EMAIL_MARKER_REPORT`.
- **12:50:03** — Re-synced template (`scripts/sync_template.py`), rebuilt both ZIPs (`Project-Local-Build.zip` and `Pending Pickup Template (macro).zip`), `test-compile.py` → COMPILE OK.
- **16:57:00** — Made root `README.md` very detailed and informative for the GitHub main page (what it does, at-a-glance table, quick start, template flow, email templates, data contract, comparison logic, output specification, aging definition, maintainer workflow, distribution, screenshots: Dashboard + one manager tab + Raw Data, timeout/validation policy, known issues, decisions/open questions, documentation links).
- **16:55:00** — Restored detailed content to root `README.md` (quick start, key artifacts, maintainer workflow, distribution, screenshots: Dashboard + one manager tab + Raw Data) so GitHub landing page shows the full picture again; root README now points to `documentation/README.md` for the comprehensive guide.
- **16:50:00** — Overhauled `README.md` (root landing page: quick start, key artifacts, maintainer workflow, distribution) and `documentation/README.md` (comprehensive user/maintainer guide: project files, repository layout, requirements, quick start, locked-down work PC rebuild, template flow, email templates, data contract, comparison logic, output specification, aging definition, maintainer workflow, screenshots, timeout/validation policy, known issues, open questions). Root README now points to `documentation/README.md` per Markdown-location rule.
- **16:50:01** — `scripts/run-report.py` now pushes all run artifacts (screenshots, zips, templates, logs) to `origin/main` after every successful test run via `git add -A` / `git commit` / `git push`, so GitHub stays in sync with each run.
- **16:45:00** — `scripts/run-report.py` rebuilt both build artifacts (`Project-Local-Build.zip` and `Pending Pickup Template (macro).zip`) via `sync_template.py` after each successful report run, then committed and pushed to GitHub.
- **16:43:00** — `Pending Pickup Template (macro).zip` reduced to contain only `Pending Pickup Template.xlsm` (`.bas` modules already embedded); `Project-Local-Build.zip` retains the no-macro `.xlsx` + all `.bas` modules + `BUILD-INSTRUCTIONS.txt`. Both built in tandem by `scripts/sync_template.py`.
- **16:40:00** — `scripts/sync_template.py` now builds `Pending Pickup Template (macro).zip` alongside `Project-Local-Build.zip` in the sync/build step.
- **16:30:00** — Overhauled KPI coloring in `vba/modFormatting.bas`: replaced near-identical light blue/green fills with a distinct palette (navy, steel blue, green, olive, gray, purple, red, dark red, maroon, gold) with white card text on both dashboard cell-cards and manager-tab shape-cards; gap fill around shape cards changed to light gray.
- **16:25:00** — Moved VBA source of truth from project root to `vba/` (7 `.bas` modules detected as Git renames); updated `scripts/sync_template.py` default `bas_dir` and build-package glob; updated `AGENTS.md`.
- **16:20:00** — Cleared all runtime logs, re-ran `run-report.py` (exit 0), regenerated all 13 report screenshots, `check-logs.py --latest` → 0 errors / 0 warnings; committed and pushed.
- **16:00:00** — Timeout policy hardened: screenshots/exports required with 120s cap (Excel killed only if >120s); non-screenshot steps 30s cap; skips flagged as warnings/errors in `check-logs.py`, `run-report.py`, and `pipeline.py`; `AGENTS.md` updated.

## 2026-11-17

- **11:46:00** — Overhauled `README.md` (current workflow, email template config/`{{REPORT}}`/Alt+Enter, template alignment for macro vs no-macro/live, scripts). Updated `.gitignore` to exclude Python cache (`__pycache__/`, `*.pyc`).

- **11:44:00** — `scripts/run-report.py` now checks Manager/Director email template fields (TO/CC/SUBJECT/BODY) and runs `sync_template.py` before report runs when macro/no-macro templates differ; updated docs to reflect run-time auto-sync plus optional save-time watcher.

- **11:36:00** — Added `scripts/watch-template-sync.py` to automatically detect stable saves to `Pending Pickup Template.xlsm`, compare Manager/Director email template fields (TO/CC/SUBJECT/BODY) against `Pending Pickup Template (no-macros).xlsx`, and run `sync_template.py` to keep macro-enabled and no-macro/ZIP artifacts aligned. Documented watcher usage in `README.md` and `documentation/README.md`.

- **11:31:00** — Email template tabs: added spacing between TO/CC/SUBJECT/BODY sections (TO `D5`, CC `D7`, SUBJECT `D9`, BODY `D11:D31`) with matching VBA/config mappings and layout migration; synced templates/build artifacts and refreshed screenshots.

- **11:14:00** — Email template tabs: BODY now uses `{{REPORT}}` marker for pasted Excel data (text above intro, below closing; if removed report pastes after body with default closing). Updated on-sheet instructions (semicolon TO/CC, `{{REPORT}}`, Alt+Enter for new line in BODY), director email set to monthly with current date appended to subject, and added "Return to START HERE" button to re-hide template tabs. VBA config cells preserved (`D5/D6/D7/D9`).


## 2026-11-16

- **14:43:00** — Removed bordered/filled "board" styling from the manager-tab `TOTAL PER SALES REP` summary block in `modFormatting.bas` (`AddManagerRepSummary`): summary now uses plain text (no background/borders) to match surrounding manager-tab style.

- **14:33:00** — Extended right-edge dashboard KPI cards to include column K: `STILL PENDING` and `MISSING PICKUP DATE` now merge across `I4:K5`/`I6:K6` and `I8:K9`/`I10:K10` in `modFormatting.bas`, removing the blank cell at the right side of those cards.
- **14:33:01** — Fixed Dashboard screenshot export omitting the bottom chart: `scripts/snapshot-report.py` now expands the captured range to include embedded chart/shape bounds below the sheet `UsedRange`, so `PENDING PICKUPS BY MANAGER (CHART)` is included in exported PNGs.
- **14:33:02** — Fixed `HideUnusedCells` in `modUtils.bas` hiding rows occupied by embedded charts/shapes: now expands the last-used row/column to include `ChartObjects`/`Shapes` bottom-right bounds so the Dashboard manager chart remains visible below the table.
- **14:33:03** — Regenerated `Pending Pickup Template.xlsm`, `Template (No Macros).xlsx`, `Project-Local-Build.zip`, and refreshed `documentation/screenshots/report-dashboard.png` from the latest report output.

- **14:24:00** — Fixed dashboard edge-to-edge spacing/alignment: set uniform column widths `A:K = 13.5`, expanded title/subtitle and chart header merges to `A1:K2`/`A3:K3`/`A26:K26` (was `A1:H2`/`A3:J3`), and sized the manager chart to full `A:K` width so KPI cards, table, and chart align consistently across the dashboard.

- **13:54:00** — Wired local build package generation into `scripts/sync_template.py` so every sync/build also regenerates `Template (No Macros).xlsx` (macro-free base from synced template) and rebuilds `Project-Local-Build.zip` with the macro-free template, all VBA source modules (`*.bas`/`*.cls`/`*.frm` in project root), and `BUILD-INSTRUCTIONS.txt`. Added `--skip-package` flag for maintenance cases where package rebuild is not needed. ZIP now includes `Pending_Pickup_Macro.bas` legacy module alongside modular sources.

- **13:48:00** — Added work-PC-friendly local build package for locked-down/macro-blocked environments (Mark-of-the-Web "Microsoft has blocked macros…"). Committed `Template (No Macros).xlsx` (macro-free base, no `vbaProject.bin`), `BUILD-INSTRUCTIONS.txt` (download/unblock/import-all-modules/save-as-`.xlsm` steps, no Python/tools), and `Project-Local-Build.zip` containing the macro-free template, all VBA source modules (`modAudit.bas`, `modCore.bas`, `modDeclarations.bas`, `modDialog.bas`, `modFormatting.bas`, `modLogging.bas`, `modUtils.bas`), and instructions. Added/committed modular `.bas` source files to repo as source of truth for rebuild; `Pending_Pickup_Macro.bas` retained as legacy monolithic reference.
- **13:48:01** — Updated `README.md` and `documentation/README.md` with "Work-PC / locked-down machines" quick-start, build package contents, distribution guidance (prebuilt `.xlsm` may still be blocked; macro-free base + modules ZIP is reliable fallback), and maintenance notes to keep ZIP/modules updated with template/VBA changes. No sensitive/PHI data included.

- **01:12:00** — Fine-tuned email prototype: rewrote `EmailSelectedManagerTab` and `EmailDashboard` in `modDialog.bas` to use `Application.Caller` + shape `AlternativeText` instead of stale global `mEmailTargetSheet` (which was overwritten during manager tab creation, causing every manager email button to email the last manager's data). Added `GetWordEditor` helper with 4-second retry for Outlook Word editor availability. Changed `SendDashboardEmail` to accept `rowCount` parameter. Replaced `On Error Resume Next` with `On Error GoTo CleanFail` + `LogErr` in all email subs.
- **01:12:00** — `modCore.bas`: removed global `mEmailTargetSheet`/`mEmailTargetManager`/`mEmailTargetRowCount` assignments during button creation. Added `emailBtn.AlternativeText = ws.Name` and `dashBtn.AlternativeText = dashboard.Name` so each email button carries its own target sheet identity.
- **01:12:00** — `modDeclarations.bas`: removed unused `EMAIL_TEMP_IMAGE` constant and unused `mEmailTargetSheet`/`mEmailTargetManager`/`mEmailTargetRowCount` globals.
- **01:12:00** — `scripts/snapshot-report.py`: made `taskkill` conditional on `--no-kill` flag; increased kill delay to 3s; set `xl.Visible = True` + `xl.WindowState = Maximized` for reliable `CopyPicture`; added 1s sleep after workbook open for rendering; wrapped Excel settings reset in try/except.
- **01:12:00** — `scripts/run-report.py`: moved screenshot regeneration to after try/finally (Excel fully shut down) to avoid COM conflicts; `snapshot-report.py` now starts its own fresh Excel instance.
- **01:17:00** — Fixed screenshot capture producing blank/raw-text images: changed `CopyPicture` format from `xlBitmap` (2) to `xlPicture` (1) which preserves vector formatting (borders, fills, colors); added 3-attempt retry for intermittent COM failures; opened workbook as read-write (not ReadOnly) so Excel fully renders; increased post-open sleep to 2s and post-paste sleep to 0.5s (`scripts/snapshot-report.py:64-96,133`).
- **01:17:00** — Added error handling + logging to `BuildDashboard` and `AddDashboardCard` in `modFormatting.bas` — both now use `On Error GoTo CleanFail` + `LogErr`, and log each KPI card as it's created. Confirmed all 8 dashboard KPI cards render with correct values (Total Pending=150, New=150, Still Pending=0, Cleared=0, Aging 30+=39, 60+=24, 90+=13, Missing Dates=4).
- **01:17:00** — Added logging to `FormatManagerSheet` in `modFormatting.bas` — now logs start/done for each of the 10 manager tabs so formatting failures are visible in `macro-debug.log`.
- **08:01:00** — Restored manager sheet formatting lost during modular split. `FormatManagerSheet` in `modFormatting.bas` now applies full data formatting (`ApplyDataFormatting` for Segoe UI font, borders, centered alignment), header styling (bold, white text, dark blue background `RGB(31,78,121)`), and alternating row colors (`RGB(242,242,242)`). Previously only age warning colors and missing-date highlighting were applied — all table borders, header formatting, and alternating rows were missing.
- **01:29:00** — Removed duplicate representative name audit entirely. Duplicate rep names are normal business behavior — one rep handles multiple pending pickups for different customers, so it is not a data quality issue. Removed `mAuditDuplicateNames` counter, `DUPLICATE_NAME` audit events, and `Duplicate Rep Names` audit sheet row from `modCore.bas`, `modAudit.bas`, and `modDeclarations.bas`. `mRepSeenDict` retained for potential future use but no longer flags duplicates. Verified: 0 DUPLICATE_NAME entries in latest run (was 104+).
- **01:25:00** — Fixed `DUPLICATE_NAME` log flooding: `modCore.bas` now logs only the first 5 duplicate names individually, then writes a single `DUPLICATE_NAME_SUMMARY` entry with total count. Previously 104 individual entries per run drowned out the log. Verified: 6 entries instead of 104 in the latest run.
- **01:25:00** — Log gap audit findings: (1) `CleanFail` does not write an `AUDIT` event on macro failure — only `LogMsg` + `LogErr`. (2) Email click handlers are never exercised in unattended runs — no end-to-end email log trail exists yet. (3) VBA `On Error Resume Next` blocks can swallow errors that Python never sees — outcome only reflects Python-level success/failure, not internal VBA swallowed errors.

- **07:49:00** — Split monolithic `Pending_Pickup_Macro.bas` (~1666 lines) into 7 modular `.bas` files for maintainability: `modDeclarations.bas` (Public constants + shared state), `modLogging.bas` (LogMsg/LogErr/BuildRunLogPath/WriteCompletionMarker), `modAudit.bas` (AuditInit/AuditLogEvent/AuditWriteSheet), `modCore.bas` (CreatePendingPickupReport main builder), `modFormatting.bas` (BuildDashboard/FormatManagerSheet), `modUtils.bas` (all utility functions), `modDialog.bas` (file dialogs + email). All modules compile cleanly (`COMPILE OK`).
- **07:49:01** — Added email feature: "Email <Manager>" button on each manager tab and "Email Dashboard to Director" button on Dashboard. Uses Outlook COM (`CreateObject("Outlook.Application")`) with inline `cid:` image snapshot of the sheet (`CopyPicture` → `ChartObject` → `PNG`). Opens Outlook inspector (`.Display`) so user types To: and hits Send — no `.oft` template file needed.
- **07:49:02** — Added email configuration cells directly on the template's START HERE sheet: CC recipients (`D16`), Subject (`D17`), Body (`D18`). Emails read from these cells so users can edit CC/subject/body right in the template; changes persist across report runs. No registry or external files required.
- **07:49:03** — Changed email body from embedded PNG screenshot (`cid:` image + `Attachments.Add`) to direct copy/paste of the Excel range into Outlook's Word editor, preserving formatting and spacing. `SendManagerEmail` and `SendDashboardEmail` now copy the sheet range to clipboard, open the Outlook inspector, and paste into `mail.GetInspector.WordEditor`. Removed `EMAIL_TEMP_IMAGE` temp file and `CopySheetSnapshot` usage from the email flow.
- **07:49:04** — Added "Setup Email CC Recipients" button on START HERE (`SetupEmailCC`) that activates and selects the CC cell for editing.
- **07:49:04** — Enhanced logging: report run start time + duration (seconds), email config reads (CC/Subject/Body), email snapshot export (success/failure path), Outlook availability check, email inspector open/close, email button clicks. All logged to `macro-debug.log` via `LogMsg`.
- **07:49:05** — Fixed `BuildManagerOutputArray` array dimension bug (`lastCol + 1` needed for the extra "Pickup Age (Days)" column — was causing "Subscript out of range" Err=9). Fixed `FormatManagerSheet` parameter mismatch (was called with 8 args, declared 12). Fixed `GetRecordKey` defensive column check (returns `""` if `keyCol <= 0` or `nameCol <= 0`). Fixed `For Each r In` → `For Each rowItem In` (VBA requires Variant for `For Each` control variable). Fixed single-line `If/ElseIf` in `CalculateAgingStats` (VBA requires block If for ElseIf chains).
- **07:49:06** — Fixed dashboard KPI card nested merge bug (cards now render label + value correctly instead of merged-parent conflict).
- **07:49:07** — Migrated all PowerShell harnesses to Python: `watch-compile-error.py`, `answer-dialogs.py`, `sync_template.py`, `snapshot-report.py`, `test-compile.py`, `run-report.py`, `audit-peer.py`, plus `documentation/view-logs.py`. All `.ps1` references updated in docs; `.ps1` removed from `section_specs` in `audit-peer.py`.
- **07:49:08** — Added auto-mode for unattended runs: `log/auto-source.txt` provides the source path (skips CURRENT picker), `log/auto-skip-prior.txt` skips PRIOR picker. Added `VBAWarnings` registry + trusted location + `AutomationSecurity` settings to bypass macro security before Excel opens.
- **07:49:09** — Graceful Excel/workbook close after report generation (`wb.Close(False)`, `excel.Quit()`) — no more lingering Excel processes or open template/temp files.
- **01:00:26** — Fixed `CopyColumnFormats` hang: was copying entire columns (1M+ rows) one at a time via clipboard, causing Excel COM to hang indefinitely. Now copies header row formats in a single clipboard operation and sets column widths directly. Report now completes end-to-end in ~3 seconds (`modUtils.bas:424-433`).
- **01:00:26** — Added granular TRACE logging in `modCore.bas` after every step between Raw Data creation and SaveAs (`BuildOutputArray`, `FormatRawData`, `CopyColumnFormats`, `MatchSourceColumnWidths`, `ApplyDateFormatsToOutput`, `ApplyDateFormatByHeader`, `HideUnusedCells` loop with sheet count, `AuditWriteSheet`, `AuditLogEvent SAVE`, `SaveAs`). Previously the macro went silent between "Raw Data sheet created" and "BuildOutputPath done", making hangs impossible to diagnose.
- **01:00:26** — Added Python-side run duration tracking in `scripts/run-report.py`: logs total run duration, final outcome, and any new report files created. Also logs screenshot regeneration status (`scripts/run-report.py:482-490`).

- **01:00:26** — Auto-regenerates test screenshots after every successful report run: `run-report.py` calls `snapshot-report.py` which deletes all existing PNGs in `documentation/screenshots/` and exports fresh images from the newest report.

## 2026-11-15

- **17:45:00** — Added comprehensive audit logging to `Pending_Pickup_Macro.bas`:
  * AuditInit/AuditLogEvent/AuditWriteSheet sub-procedures for run tracking
  * Audit fields: total records, pending counts, aging breakdown, missing dates,
    duplicate keys, missing managers, error count
  * AuditInit called at report start with source file path
  * AuditLogEvent called at key points (start, counts, aging stats, save)
  * AuditWriteSheet adds an "Audit Log" sheet to the output workbook with full
    run details, record counts, data quality issues, and all audit entries
  * Audit entries written to both macro-debug.log and per-run log
- **17:05:00** — Fixed manager tab KPI card layout: 5 cards on row 4-6 and 4 cards on row 8-10 so no card sits alone (`Pending_Pickup_Macro.bas:945-955`).
- **17:00:00** — Added `HideUnusedCells` so every sheet hides all rows/columns beyond its used range, keeping the report clean and focused only on actual data (`Pending_Pickup_Macro.bas:469-475,1857-1893`).
- **16:55:00** — Improved spacing on dashboard and manager tabs: KPI card row heights 34→38, data row heights 22→28, min column width 14→16, font sizes increased for readability (`Pending_Pickup_Macro.bas:738,741,973-975,1091,1150`).
- **16:54:00** — Improved screenshot quality: `snapshot-report.py` now sets zoom to 100% before CopyPicture and uses wider charts (110px/col, 30px/row) so content isn't squished onto one page width (`scripts/snapshot-report.py:48-57`).
- **16:50:00** — Fixed screenshot capture: rewrote `scripts/snapshot-report.py` using the working PM Project technique (standalone Chart sheet, `xlBitmap` CopyPicture format 2, sleep after paste) so screenshots capture real data instead of blank/garbled output. All 12 sheets (Dashboard, 10 manager tabs, Raw Data) now export as valid PNGs (458KB-2MB each).
- **16:49:00** — Removed broken VBA screenshot export code (`ExportSheetAsPng`, `ExportAllSheetsPng`) from `Pending_Pickup_Macro.bas`; screenshots now handled purely from PowerShell via `snapshot-report.py`.
- **16:48:00** — Major root README.md overhaul: detailed GitHub landing page with app workflow (5 steps), repository layout, output specification, KPI definitions (dashboard + manager tabs), comparison logic, aging definition, requirements, distribution notes, and embedded screenshots for Dashboard, all 10 manager tabs, and Raw Data.
- **16:47:00** — Replaced 13 old broken PNGs (0-byte / garbled) in `documentation/screenshots/` with 12 new screenshots captured from a live report run using real test data, named `report-*.png`.
- **16:35:00** — Overhauled dashboard and manager tab KPI cards with descriptive labels (e.g. "CURRENT" → "CURRENT PERIOD", "30+ DAYS" → "AGING 30+ DAYS", "AVG AGE" → "AVERAGE PICKUP AGE", "% 30+" → "% OVER 30 DAYS") and a cohesive color palette (blue for counts, red gradient for aging, amber for missing dates, green for new, purple for %) (`Pending_Pickup_Macro.bas:627-641,939-948`).
- **16:34:00** — Improved spacing on manager tabs: KPI card row heights increased from 26 to 34, data row heights from 22 to 24, min column width from 12 to 14, KPI card font sizes increased (label 9→10, value 20→22) (`Pending_Pickup_Macro.bas:881-891,973-975,1091,1150`).
- **16:33:00** — Improved dashboard spacing: KPI card row heights from 30 to 34, data row heights from 22 to 24 (`Pending_Pickup_Macro.bas:738,741`).
- **16:25:00** — Aligned dashboard KPI cards on equal column widths (A:J) so cards line up vertically with each other and with the manager table; removed the "PENDING BY EQUIPMENT TYPE" chart section since the report tracks pickup status, not inventory (`Pending_Pickup_Macro.bas:627-642,826-900`).
- **16:24:00** — Cleaned up unused variable declarations (`eqTitleRow`, `eqRow`, `eqCol`, `eqAgg`, `eqKeys`, `eqName`, `aAgg`, `bAgg`, `sTmp`) left after removing the equipment Pareto section (`Pending_Pickup_Macro.bas:537-542`).
- **19:15:00** — Added AGENTS.md rule (section 9) requiring all existing PNGs in `documentation/screenshots/` to be deleted before each test run so screenshots always reflect the latest report output, never a stale mix.
- **19:14:00** — Updated `scripts/snapshot-report.py` to delete every existing PNG in `documentation/screenshots/` before regenerating them from the current report run.
- **19:03:00** — Verified compilation via Excel COM automation: `Pending_Pickup_Macro.bas` imports into a fresh `PendingPickupReport` module with no syntax or compile errors.
- **19:02:00** — Fixed a latent subscript error in `SaveFolderSetting`: `Left$(..., InStrRev(...)-1)` errored when the path had no separator; now guards `sep > 0` (`Pending_Pickup_Macro.bas:1820-1822`).
- **19:01:00** — Fixed the Dashboard "% CHANGE" KPI card not applying its `0.0%` number format: the value was written via `CStr` (text), so it showed a raw decimal. It now writes the numeric `pctChange` into the value cell so it renders as a true percentage (`Pending_Pickup_Macro.bas:627-631`).
- **19:00:00** — Fixed state leak in `CreatePendingPickupReport`: early exit when `EnsureMacroEnabledWorkbook()` fails now `GoTo CleanExit` instead of `Exit Function`, so `mInReportRun`, `mRunLogPath`, and Excel app state (ScreenUpdating, Calculation, Alerts, Events) are properly reset (`Pending_Pickup_Macro.bas:185`).
- **18:49:00** — Documented the `audit-peer` workflow in `documentation/AGENTS.md`: require read-only cross-project audits, explicit approval before porting, separate agent-rule review, full validation after any port, and Markdown reports under `documentation/`; corrected the audit's rules section to read `documentation/AGENTS.md`.
- **18:12:00** — Moved `AGENTS.md` into `documentation/` and updated the repository rule so all Markdown files, including agent instructions, are stored there rather than the project root.

- **18:08:01** — Standardized every runtime log entry on full `yyyy-MM-dd HH:mm:ss` timestamps and added `documentation/view-logs.py` to reorder every runtime `.log` and `.txt` file newest first with dated headers.
- **18:08:00** — Moved `README.md`, `changelog.md`, and both `aoti-logo-2023` source images into `documentation/`; updated documentation references and added the documentation-location rule for Markdown files.
- **17:44:00** — Moved operational PowerShell harnesses into `scripts/` and changed their default paths to derive the project root from `$PSScriptRoot`; updated README and agent commands to use the new locations.
- **17:48:00** — Removed the final report-completion `vbYesNo` prompt so unattended report runs exit after saving instead of hanging at “Report complete! Open the finished report now?”.
- **17:41:00** — Restored `Reports/` and `log/` as tracked project directories using `.gitkeep`. Generated reports and prior diagnostic log contents remain excluded as runtime artifacts; new logs and reports will be written into these folders when the harness or macro runs.
- **17:39:00** — Removed generated `Reports/` workbooks, runtime `log/` files, `screen.png`, `_run_temp.xlsm`, and `probe_ran.txt` from the project directory. Updated `.gitignore` to keep generated report artifacts and runtime diagnostics out of version control.
- **05:25:00** — Fixed "Type mismatch" runtime error in `ApplyAgeWarningFont` caused by `CDbl(v)` on formula
  strings (`=IF(...)`) in the Pickup Age column. Wrapped with `On Error Resume Next`.
- **05:25:01** — Added VBA-level logging to `macro-debug.log` with timestamps and procedure names for all
  key steps: `CreatePendingPickupReport`, `FormatManagerSheet`, `BuildDashboard`, etc.
- **05:25:02** — Added `LogMsg` and `LogErr` helper subs; all error handlers now log to macro-debug.log
  with procedure name, error description, error number, and line number.
- **05:25:03** — Fixed harness cleanup order: Excel is killed first in the `finally` block before releasing
  COM objects or removing temp files; removed screenshot delay.
- **05:25:04** — Added comprehensive error handling (`On Error Resume Next` + bounds checking) to all
  data-processing functions: `CalculateAgingStats`, `CalculateAverageAge`, `CountNewManagerRows`,
  `GetRecordKey`, `FindHeaderColumnFromArray`, `NormalizeDateColumn`, `TryGetExcelDate`,
  `BuildOutputArray`, `BuildManagerOutputArray`, `SortManagerRowsBySalesRep`, `SortManagerColumns`,
  `SortManagersByLastName`, `SortTextArray`.
- **05:25:05** — Added comprehensive error handling to all formatting functions: `ApplyRepBanding`,
  `ApplyAgeWarningFont`, `HighlightMissingPickupDates`, `FormatYesNoColumns`, `HighlightNewRecords`.
- **05:25:06** — Added error handling to file I/O functions: `SelectExcelFile`, `BuildOutputPath` with
  error logging to macro-debug.log.
- **05:25:07** — Verified full report generation with no errors: `Pending Pickup 2026-11-13_223136.xlsx`
  (97KB) generated successfully with Dashboard, 10 manager tabs, and Raw Data.

## 2026-11-13

- **23:45:00** — Added a read-only cross-project audit that runs from either project:
  `scripts/audit-peer.py` compares this project with the sibling `PM Project` and
  reports files only one side has, files that differ (with the functions or Markdown
  sections each side adds), and VBA procedure names that are shared versus peer-only.
  `-Sections` narrows the audit, `-Report <path>` saves Markdown, `-SymbolLimit` caps
  the listing. The peer is auto-discovered from a sibling directory holding a root
  `*.bas` and a `scripts/` folder; `-PeerPath` overrides it. Added
  `.kilo/command/audit.md` so `/audit` runs it and asks for a shortlist of port
  candidates with effort and risk; the command never copies anything.
- **15:02:38** — Dashboard upgrades: "Report generated <timestamp> | Source: <file>" line, bold TOTAL row,
  per-manager Cleared and Still Pending columns, data bars on Current, a 3-colour scale on 90+ Days, and a
  "Current Pending Pickups by Manager" bar chart. Dashboard tab colour set to dark blue.
- **15:02:38** — Manager tabs: light-blue tab colour, landscape fit-to-width print setup with repeating header
  row, and red highlighting on the AUDIT: TP End Date cell when it is missing.
- **15:02:38** — Raw Data: Patient Date of Birth and Last Modified now render as mm/dd/yyyy (previously serial
  numbers), grey tab colour, and landscape fit-to-width print setup with repeating header row.
- **15:02:38** — Fixed a compile error in the dashboard colour-scale call: `AddColorScale(3)` plus
  `ColorScaleCriteria` is used instead of the unsupported `AddColorScale 2, RGB(...), RGB(...)` form.
- **14:40:01** — Fixed runtime error "Unable to set the FreezePanes property of the Window class" that
  aborted the build: the Dashboard / manager / Raw Data window-formatting calls (`ActiveWindow.FreezePanes`,
  `SplitRow`, `DisplayGridlines`, `Zoom`) are now wrapped in `On Error Resume Next` / `On Error GoTo 0` so a
  window that is not active can no longer fail the macro.
- **14:40:01** — Fixed the Dashboard logo not being embedded: `AddLogoFromTemplate` now selects `A3`, calls
  `DoEvents` after `src.Copy` so the clipboard is ready before `ws.Paste`, and runs under error suppression.
  Generated and verified a full sample report at `Reports/Pending Pickup 2026-11-13.xlsx` (Dashboard + 10
  manager tabs + Raw Data, logo embedded, all KPI totals reconcile).
- **14:04:29** — Redesigned the `START HERE` sheet into a clean branded landing page: title + logo,
  4-step how-it-works, styled rounded-rectangle CTA button, and a good-to-know panel; replaced the plain
  Forms button with an autoshape and updated `sync_template.py` and `SetupPendingPickupButton` to match.
- **13:52:11** — Embedded the company logo (`aoti-logo-2023.svg` rasterized to a transparent 476×136 PNG)
  on `START HERE` as `AOTILogo`, making the template fully self-contained.
- **13:52:11** — Added `AddLogoFromTemplate` to the builder and placed the logo at the top-right of the
  generated Dashboard (`I1:J2`); re-synced the template.
- **13:52:11** — Documented single-file distribution: end users receive only the `.xlsm`; `.bas`,
  `sync_template.py`, README, changelog, and fake data are maintainer-only.
- **13:46:24** — Added `sync_template.py` to re-import `Create_Pending_Pickup_Report.bas` into the
  template's embedded VBA and ensure the `START HERE` button; enabled `AccessVBOM=1` for Excel 16.0.
- **13:46:24** — Synced the v1.0 module into `Create_Pending_Pickup_Template.xlsm` and pre-installed the
  `CREATE PENDING PICKUP REPORT` button (OnAction = `CreatePendingPickupReport`), so end users never
  import the `.bas`.
- **13:41:14** — Added this `changelog.md`.
- **13:41:13** — Removed legacy `Pending_Pickup_Report_README_v3.5.txt`.
- **13:41:12** — Renamed template `Pending_Pickup_Report_Template_v3.5.xlsm` →
  `Create_Pending_Pickup_Template.xlsm`; renamed source `Create_Pending_Pickup_Report_v3.5.bas` →
  `Create_Pending_Pickup_Report.bas`.
- **13:41:11** — Set version to `1.0` across `README.md`, the VBA source (`MACRO_VERSION = 1.0`), and the
  template `START HERE`; retired the in-module per-version changelog.
- **13:41:10** — Added prior-month fixture `Reports/Pending Pickup 2026-10-13.xlsx` (126 rows) for
  comparison testing; expected result against the fake export is ~114 still pending / ~36 new / ~12 cleared.
- **13:41:09** — Created `Reports/` output folder and pre-set the macro's remembered `OutputFolder`
  (registry `HKCU\Software\VB and VBA Program Settings\PendingPickupReport\Settings\OutputFolder`).
- **13:41:08** — Generated fake data in `AllOpenTreatmentPlansResults618.xlsx`: 150 Pending Pickup rows +
  15 other-status rows, 10 managers, 48 unique reps (no rep under two managers), realistic aging spread,
  4 missing `AUDIT: TP End Date`, 43 blank `EMPI`.
- **13:41:07** — Overhauled `README.md` to reflect current scope: confirmed business context, data
  contract, comparison logic, output specification (Dashboard / manager tabs / Raw Data), aging
  definition, known issues, and decisions/open questions.