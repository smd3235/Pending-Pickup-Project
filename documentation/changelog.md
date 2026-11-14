# Changelog

All notable changes to the Pending Pickup Report Builder. Each entry is stamped
`YYYY-MM-DD HH:MM:SS` (local time). Version `1.0` is the current baseline.

## 2026-11-14

- **19:03:00** — Verified compilation via Excel COM automation: `Pending_Pickup_Macro.bas` imports into a fresh `PendingPickupReport` module with no syntax or compile errors.
- **19:02:00** — Fixed a latent subscript error in `SaveFolderSetting`: `Left$(..., InStrRev(...)-1)` errored when the path had no separator; now guards `sep > 0` (`Pending_Pickup_Macro.bas:1820-1822`).
- **19:01:00** — Fixed the Dashboard "% CHANGE" KPI card not applying its `0.0%` number format: the value was written via `CStr` (text), so it showed a raw decimal. It now writes the numeric `pctChange` into the value cell so it renders as a true percentage (`Pending_Pickup_Macro.bas:627-631`).
- **19:00:00** — Fixed state leak in `CreatePendingPickupReport`: early exit when `EnsureMacroEnabledWorkbook()` fails now `GoTo CleanExit` instead of `Exit Function`, so `mInReportRun`, `mRunLogPath`, and Excel app state (ScreenUpdating, Calculation, Alerts, Events) are properly reset (`Pending_Pickup_Macro.bas:185`).
- **18:49:00** — Documented the `audit-peer` workflow in `documentation/AGENTS.md`: require read-only cross-project audits, explicit approval before porting, separate agent-rule review, full validation after any port, and Markdown reports under `documentation/`; corrected the audit's rules section to read `documentation/AGENTS.md`.
- **18:12:00** — Moved `AGENTS.md` into `documentation/` and updated the repository rule so all Markdown files, including agent instructions, are stored there rather than the project root.

- **18:08:01** — Standardized every runtime log entry on full `yyyy-MM-dd HH:mm:ss` timestamps and added `documentation/view-logs.ps1` to reorder every runtime `.log` and `.txt` file newest first with dated headers.
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
  `scripts/audit-peer.ps1` compares this project with the sibling `PM Project` and
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
  Forms button with an autoshape and updated `sync_template.ps1` and `SetupPendingPickupButton` to match.
- **13:52:11** — Embedded the company logo (`aoti-logo-2023.svg` rasterized to a transparent 476×136 PNG)
  on `START HERE` as `AOTILogo`, making the template fully self-contained.
- **13:52:11** — Added `AddLogoFromTemplate` to the builder and placed the logo at the top-right of the
  generated Dashboard (`I1:J2`); re-synced the template.
- **13:52:11** — Documented single-file distribution: end users receive only the `.xlsm`; `.bas`,
  `sync_template.ps1`, README, changelog, and fake data are maintainer-only.
- **13:46:24** — Added `sync_template.ps1` to re-import `Create_Pending_Pickup_Report.bas` into the
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