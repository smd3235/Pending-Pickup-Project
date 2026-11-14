# Pending Pickup Report Builder

This repository builds a monthly **Pending Pickup** report from an All Open Treatment Plans export — a KPI dashboard, one tab per Assigned Sales Manager, and filtered raw data.

## Where the docs live

All documentation (README, changelog, agent rules, screenshots) is maintained in `documentation/`.

- **User guide:** `documentation/README.md` — full setup, usage, data contract, output spec, known issues, and screenshots.
- **Change log:** `documentation/changelog.md`
- **Agent rules:** `documentation/AGENTS.md`
- **QA screenshots:** `documentation/screenshots/`

## Quick start

End users open `Pending Pickup Template.xlsm`, click **CREATE PENDING PICKUP REPORT** on `START HERE`, and follow the prompts.

Maintainers edit `Pending_Pickup_Macro.bas`, then run `scripts\sync_template.ps1` to push changes into the template.

## Repository layout

| Path | Purpose |
|------|---------|
| `Pending Pickup Template.xlsm` | Master workbook users run (embedded VBA + button) |
| `Pending_Pickup_Macro.bas` | VBA source of truth |
| `scripts\sync_template.ps1` | Re-import `.bas` into the template |
| `scripts\snapshot-report.ps1` | Export report sheets as PNG screenshots |
| `Reports/` | Generated report output (not committed) |
| `log/` | Runtime logs (not committed) |
| `documentation/` | All Markdown docs and screenshots |
| `.kilo/` | Kilo agent/command configuration |

## License

Internal tooling. See `documentation/README.md` for distribution notes (end users receive only the `.xlsm`).
