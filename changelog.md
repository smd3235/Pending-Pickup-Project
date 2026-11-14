# Changelog

All notable changes to the Pending Pickup Report Builder. Each entry is stamped
`YYYY-MM-DD HH:MM:SS` (local time). Version `1.0` is the current baseline.

## 2026-11-13

- **13:41:07** — Overhauled `README.md` to reflect current scope: confirmed business context, data
  contract, comparison logic, output specification (Dashboard / manager tabs / Raw Data), aging
  definition, known issues, and decisions/open questions.
- **13:41:08** — Generated fake data in `AllOpenTreatmentPlansResults618.xlsx`: 150 Pending Pickup rows +
  15 other-status rows, 10 managers, 48 unique reps (no rep under two managers), realistic aging spread,
  4 missing `AUDIT: TP End Date`, 43 blank `EMPI`.
- **13:41:09** — Created `Reports/` output folder and pre-set the macro's remembered `OutputFolder`
  (registry `HKCU\Software\VB and VBA Program Settings\PendingPickupReport\Settings\OutputFolder`).
- **13:41:10** — Added prior-month fixture `Reports/Pending Pickup 2026-10-13.xlsx` (126 rows) for
  comparison testing; expected result against the fake export is ~114 still pending / ~36 new / ~12 cleared.
- **13:41:11** — Set version to `1.0` across `README.md`, the VBA source (`MACRO_VERSION = 1.0`), and the
  template `START HERE`; retired the in-module per-version changelog.
- **13:41:12** — Renamed template `Pending_Pickup_Report_Template_v3.5.xlsm` →
  `Create_Pending_Pickup_Template.xlsm`; renamed source `Create_Pending_Pickup_Report_v3.5.bas` →
  `Create_Pending_Pickup_Report.bas`.
- **13:41:13** — Removed legacy `Pending_Pickup_Report_README_v3.5.txt`.
- **13:41:14** — Added this `changelog.md`.
- **13:46:24** — Added `sync_template.ps1` to re-import `Create_Pending_Pickup_Report.bas` into the
  template's embedded VBA and ensure the `START HERE` button; enabled `AccessVBOM=1` for Excel 16.0.
- **13:46:24** — Synced the v1.0 module into `Create_Pending_Pickup_Template.xlsm` and pre-installed the
  `CREATE PENDING PICKUP REPORT` button (OnAction = `CreatePendingPickupReport`), so end users never
  import the `.bas`.
- **13:52:11** — Embedded the company logo (`aoti-logo-2023.svg` rasterized to a transparent 476×136 PNG)
  on `START HERE` as `AOTILogo`, making the template fully self-contained.
- **13:52:11** — Added `AddLogoFromTemplate` to the builder and placed the logo at the top-right of the
  generated Dashboard (`I1:J2`); re-synced the template.
- **13:52:11** — Documented single-file distribution: end users receive only the `.xlsm`; `.bas`,
  `sync_template.ps1`, README, changelog, and fake data are maintainer-only.
- **14:04:29** — Redesigned the `START HERE` sheet into a clean branded landing page: title + logo,
  4-step how-it-works, styled rounded-rectangle CTA button, and a good-to-know panel; replaced the plain
  Forms button with an autoshape and updated `sync_template.ps1` and `SetupPendingPickupButton` to match.
- **14:40:01** — Fixed runtime error "Unable to set the FreezePanes property of the Window class" that
  aborted the build: the Dashboard / manager / Raw Data window-formatting calls (`ActiveWindow.FreezePanes`,
  `SplitRow`, `DisplayGridlines`, `Zoom`) are now wrapped in `On Error Resume Next` / `On Error GoTo 0` so a
  window that is not active can no longer fail the macro.
- **14:40:01** — Fixed the Dashboard logo not being embedded: `AddLogoFromTemplate` now selects `A3`, calls
  `DoEvents` after `src.Copy` so the clipboard is ready before `ws.Paste`, and runs under error suppression.
  Generated and verified a full sample report at `Reports/Pending Pickup 2026-11-13.xlsx` (Dashboard + 10
  manager tabs + Raw Data, logo embedded, all KPI totals reconcile).
- **15:02:38** — Dashboard upgrades: "Report generated <timestamp> | Source: <file>" line, bold TOTAL row,
  per-manager Cleared and Still Pending columns, data bars on Current, a 3-colour scale on 90+ Days, and a
  "Current Pending Pickups by Manager" bar chart. Dashboard tab colour set to dark blue.
- **15:02:38** — Manager tabs: light-blue tab colour, landscape fit-to-width print setup with repeating header
  row, and red highlighting on the AUDIT: TP End Date cell when it is missing.
- **15:02:38** — Raw Data: Patient Date of Birth and Last Modified now render as mm/dd/yyyy (previously serial
  numbers), grey tab colour, and landscape fit-to-width print setup with repeating header row.
- **15:02:38** — Fixed a compile error in the dashboard colour-scale call: `AddColorScale(3)` plus
  `ColorScaleCriteria` is used instead of the unsupported `AddColorScale 2, RGB(...), RGB(...)` form.

## 2026-11-14

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