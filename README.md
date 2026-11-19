# Pending Pickup Report Builder

Monthly Pending Pickup worklist + KPI dashboard builder for Excel. Converts a raw **All Open Treatment Plans** export into a formatted report: a dashboard with KPI cards and a manager breakdown, one tab per Assigned Sales Manager with per-rep detail, filtered Raw Data, an Audit Log, and optional comparison against the prior month's report.

## Documentation

Full user/maintainer guide: [`documentation/README.md`](documentation/README.md)

Agent/project rules (maintainers): [`documentation/AGENTS.md`](documentation/AGENTS.md)

Dated change log: [`documentation/changelog.md`](documentation/changelog.md)

## Quick start

1. Open `Pending Pickup Template.xlsm` (enable macros / trust the VBA project).
2. On **START HERE**, click **CREATE PENDING PICKUP REPORT**.
3. Select the current All Open Treatment Plans export.
4. (Optional) Select last month's report for comparison, or **Cancel** to skip.
5. Choose where to save the report (remembered via registry `SaveSetting`).

The report opens on **Dashboard** and is saved as `Pending Pickup yyyy-mm-dd_HHMMSS.xlsx`.

## Distribution

End users receive only `Pending Pickup Template.xlsm`. For locked-down work PCs where downloaded `.xlsm` files are blocked, extract `Project-Local-Build.zip`, import all seven `vba/*.bas` modules into the macro-free `.xlsx`, and save as `.xlsm`. See [`BUILD-INSTRUCTIONS.txt`](BUILD-INSTRUCTIONS.txt).
