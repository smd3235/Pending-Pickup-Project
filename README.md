# Pending Pickup Report Builder

An Excel macro-enabled workbook (`.xlsm`) that turns a raw **All Open Treatment Plans** export into a formatted **Pending Pickup** report: a KPI dashboard, one tab per Assigned Sales Manager, and filtered raw data. It can also compare the current run against the prior month's Pending Pickup report to show what is new, what cleared, and what is still pending.

---

## What this does

Every month, the report pulls **every patient in the country** from the All Open Treatment Plans export. It references any patient who has equipment **pending pickup**, and the resulting report is sent to each manager covering their team/region — so each manager sees exactly the patients they need to close out and pick up equipment for.

- **Pending Pickup is the only treatment status included.** A record leaves the report ("cleared") when the equipment is picked up and the patient is closed out, **or** when the patient is returned to active and the status is no longer Pending Pickup. Both are legitimate exits, so "in prior but not current" means cleared.
- **Pickup Age** = days from `AUDIT: TP End Date` to today. **Older is worse** — it means the account has been waiting longer for pickup.
- **Missing `AUDIT: TP End Date`** means whoever flagged the account for pickup forgot to enter the date, so the wait time is unknown; these are surfaced as `MISSING DATE`.
- **Cadence:** run monthly. Each run compares against the **prior month's generated Pending Pickup report** to show progress (cleared / still pending) plus what is new.
- **EMPI** is present in the export but is **not shown in the final report** — it is treated as noise. It is still used internally as the comparison key, with `Name` as the fallback.
- **Roles:** `Assigned Sales Manager` owns a manager tab; `Sales Representative` is the rep on the record. No rep works for more than one manager.

---

## Repository layout

| Path | Purpose |
|------|---------|
| `Pending Pickup Template.xlsm` | Master workbook users run (embedded VBA + `START HERE` button + company logo) |
| `Pending_Pickup_Macro.bas` | Readable VBA source of truth for the builder |
| `scripts\sync_template.ps1` | Re-imports `.bas` into the template's embedded VBA and ensures the `START HERE` button |
| `scripts\snapshot-report.ps1` | QA script: exports every sheet of a generated report as PNG into `documentation/screenshots/` |
| `Reports/` | Output folder for finished `.xlsx` reports (build artifact, not committed) |
| `log/` | Runtime logs (`macro-debug.log`, etc.) — runtime artifact, not committed |
| `documentation/` | All Markdown documentation, changelog, agent rules, and QA screenshots |
| `.kilo/` | Kilo agent/command configuration |

---

## Quick start

### End users (no setup)

1. Open `Pending Pickup Template.xlsm`.
2. Click **CREATE PENDING PICKUP REPORT** on `START HERE`.
3. Select the **current** All Open Treatment Plans export when prompted.
4. Select the **prior** Pending Pickup report for comparison (or Cancel for no comparison).
5. The report is generated: Dashboard, one tab per manager, Raw Data.
6. Choose where to save it. The save location is remembered for next time.

### Maintainers

After editing `Pending_Pickup_Macro.bas`, re-sync the template:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\sync_template.ps1
```

This requires desktop Excel with "Trust access to the VBA project object model" enabled (`HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1`).

To refresh QA screenshots after a successful run:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\snapshot-report.ps1
```

---

## Output specification

Each generated `.xlsx` contains three areas:

### Dashboard
- Title banner with company logo.
- **KPI cards (5 × 2 grid):** CURRENT, PRIOR, NEW, CLEARED, STILL PENDING on row 1; % CHANGE, 30+ DAYS, 60+ DAYS, 90+ DAYS, MISSING DATE on row 2.
- **Pending Pickups by Manager table:** manager name (hyperlink to tab), Current, New, 30+ Days, 60+ Days, 90+ Days, Missing Date, Avg Age, % 30+.
- Aging breakdown table and equipment Pareto chart.

### Manager tabs (one per Assigned Sales Manager)
- Manager name banner at top.
- KPI cards: PENDING PICKUPS, SALES REPS, NEW, 30+ DAYS, 60+ DAYS, 90+ DAYS, MISSING DATE.
- Data table sorted by Sales Representative, then Name; rows banded by rep; new records bold + green fill; Pickup Age > 30 in red text.
- Freeze panes at the header row; gridlines off; zoom 85%.

### Raw Data tab
- Current pending rows only, original column order, filterable table with frozen header.

---

## Comparison logic

- A row is "pending" when `Treatment Status` contains `Pending Pickup` (case-insensitive).
- Key = `EMPI` (trimmed, case-insensitive); if blank, fall back to `Name`; if both blank, `ROW-<n>`.
- **Still Pending** = key in both current and prior.
- **New** = key in current but not prior.
- **Cleared / No Longer Pending** = key in prior but not current.

---

## Aging definition

- **Pickup Age (Days)** = today − `AUDIT: TP End Date`, floored at 0 for future dates.
- **30+/60+/90+** buckets count rows with age **strictly greater than** 30 / 60 / 90 days.
- **Missing Date** counts pending rows with no parseable `AUDIT: TP End Date`.

---

## Screenshots

QA screenshots of the generated report (Dashboard KPI cards, manager tabs, Raw Data) are stored in `documentation/screenshots/` and embedded in `documentation/README.md`.

---

## Known issues / technical debt

1. **Embedded VBA vs `.bas` source.** The `.xlsm` embeds its own compiled copy; editing the `.bas` does **not** update the template. After every `.bas` change, run `sync_template.ps1`.
2. **Vestigial `_Config` sheet.** All labels are hardcoded `Const` values; the macro never reads `_Config`.
3. **Aging boundary.** KPI labels say "30+ / 60+ / 90+ days" but the code uses `> 30 / > 60 / > 90`.
4. **No automated tests.** The builder is exercised manually with the bundled fake export and prior fixture.

---

## Distribution

End users receive only `Pending Pickup Template.xlsm`. It is self-contained (macro, button, and company logo embedded). The `.bas`, PowerShell scripts, documentation, screenshots, and fake data are maintainer-only.

---

## Full documentation

The detailed user guide, changelog, agent rules, and full known-issues list live in `documentation/`:

- [`documentation/README.md`](documentation/README.md) — complete user guide, data contract, output spec, known issues, decisions & open questions
- [`documentation/changelog.md`](documentation/changelog.md) — dated change log (newest to oldest)
- [`documentation/AGENTS.md`](documentation/AGENTS.md) — agent operating rules for this repository
