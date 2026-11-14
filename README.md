# Pending Pickup Report Builder

This repository builds a monthly **Pending Pickup** report from an All Open Treatment Plans export — a KPI dashboard, one tab per Assigned Sales Manager, and filtered raw data.

## Full documentation

All documentation (user guide, setup, data contract, output spec, known issues, screenshots, changelog, agent rules) is maintained under `documentation/`, per repository convention.

- [**User guide** → `documentation/README.md`](documentation/README.md) — complete setup, usage, output specification, comparison logic, aging definition, known issues, and embedded QA screenshots.
- [**Changelog** → `documentation/changelog.md`](documentation/changelog.md)
- [**Agent rules** → `documentation/AGENTS.md`](documentation/AGENTS.md)

## Quick start

End users open `Pending Pickup Template.xlsm`, click **CREATE PENDING PICKUP REPORT** on `START HERE`, and follow the prompts.

Maintainers edit `Pending_Pickup_Macro.bas`, then run:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\sync_template.ps1
```

## Repository layout

| Path | Purpose |
|------|---------|
| `Pending Pickup Template.xlsm` | Master workbook users run (embedded VBA + button + logo) |
| `Pending_Pickup_Macro.bas` | VBA source of truth |
| `scripts\sync_template.ps1` | Re-import `.bas` into the template |
| `scripts\snapshot-report.ps1` | Export report sheets as PNG screenshots |
| `Reports/` | Generated report output (not committed) |
| `log/` | Runtime logs (not committed) |
| `documentation/` | All Markdown docs and QA screenshots |
| `.kilo/` | Kilo agent/command configuration |

## License

Internal tooling. End users receive only `Pending Pickup Template.xlsm`; the `.bas`, scripts, documentation, screenshots, and fake data are maintainer-only.
