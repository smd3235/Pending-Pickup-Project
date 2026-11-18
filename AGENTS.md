# Pending Pickup Project — Agent Instructions

## Hard rule: 30-second timeout (no hangs)

- **Everything other than screenshot/export operations must complete within a few seconds; if any such operation takes longer than 30 seconds, treat it as a failure/hang.** This applies to Excel/COM calls, macro/report runs, compile/sync, subprocesses, builds, log scan/validation steps, and script waits.
- **On timeout or stall (non-screenshot work):** stop/kill the process (especially Excel), read the relevant logs, identify the cause, apply a fix, then re-run. Do not increase waits to “wait it out”.
- **Exception — screenshots/exports:** screenshot/export/regeneration (e.g. `scripts/snapshot-report.py`, report screenshot steps) may take longer than 30 seconds. **Do not close or kill Excel/the app while screenshots/exports are running.** Let the capture complete; log duration and investigate if unexpectedly slow, but never interrupt a running screenshot/export by killing the app.
- Use `scripts/excel_guard.py` (`DEFAULT_TIMEOUT=30`, `ExcelGuard`, `DialogWatcher`, `kill_excel_processes`) for guarded Excel COM work. Use `scripts/pipeline.py` for end-to-end runs with a 30s cap on non-screenshot steps; screenshot steps run without the 30s kill cap.
- Python subprocess calls for non-screenshot work must have an explicit `timeout <= 30` (or be bounded/monitored). Avoid fixed sleeps that push a non-screenshot operation over 30s; prefer short waits + retries with an overall cap.
- Log hangs/timeouts to `log/excel-hangs.log` and step results to `log/pipeline.log`; inspect `log/compile-test.log`, `log/watcher.log`, `log/answer-dialogs.log`, and `log/pending-pickup_*.log` when diagnosing.

## Project layout and conventions

- **VBA source of truth:** `.bas` modules in `vba/` (`modCore.bas`, `modDialog.bas`, `modFormatting.bas`, `modLogging.bas`, `modAudit.bas`, `modDeclarations.bas`, `modUtils.bas`, etc.). Do not hand-edit generated artifacts except via sync/build flow.
- **Templates/artifacts:** `Pending Pickup Template.xlsm` (macro/test), `Pending Pickup Template (no-macros).xlsx`, `Project-Local-Build.zip` (work-PC: no-macro `.xlsx` + `.bas` modules + instructions) and `Pending Pickup Template (macro).zip` (macro-enabled `.xlsm` only, created in tandem; `.bas` already embedded). Keep aligned via `scripts/sync_template.py`.
- **Logs:** `log/` (lowercase) — `pending-pickup_*.log` (runner and/or VBA per-run), `macro-debug.log`, `pending-pickup-complete.txt`, `compile-test.log`, `watcher.log`, `answer-dialogs.log`, `excel-hangs.log`, `pipeline.log`.
- **Reports/output:** `Reports/` (generated `.xlsx`, ignore `~$` lock files). **Screenshots/docs:** `documentation/screenshots/`, `documentation/changelog.md`.
- **Config:** prefer `.kilo/command/*.md` and `.kilo/agent/*.md`; do not use `.kilocode/`.

## Standard workflow (mandatory loop for code/template changes)

1. **Compile gate:** `python scripts/test-compile.py` — VBA must compile cleanly (probe runs). Dialogs captured via `scripts/watch-compile-error.py`.
2. **Sync/build:** `python scripts/sync_template.py` — sync modules/buttons/templates into `.xlsm`, rebuild no-macro `.xlsx`, `Project-Local-Build.zip`, and `Pending Pickup Template (macro).zip`.
3. **Run report:** `python scripts/run-report.py` — end-to-end run (auto email-config sync check, unattended dialogs/skip-prior as configured), writes run logs/marker and regenerates report screenshots on success. After a successful run, rebuilds `Project-Local-Build.zip` and `Pending Pickup Template (macro).zip` via `sync_template.py`. Screenshots must complete; if regeneration exceeds 120s the app/Excel is killed and logs must be checked before rerunning.
4. **Scan/validate logs and output:** `python scripts/check-logs.py --latest`, `python scripts/monitor-logs.py`, `python scripts/analyze-trace.py`, `python scripts/validate-report.py`.
5. **Code audit (when changing scripts/code):** `python scripts/audit-code.py`.
6. **Pipeline:** `python scripts/pipeline.py` runs compile→sync→run→screenshots→scan/validate (and `audit-code` if present). Non-screenshot steps are capped at 30s (Excel killed on timeout/failure with log-tail diagnosis). Screenshot/export steps are required and capped at 120s (Excel killed only if >120s; allowed to finish under 2 minutes). Use `--skip-snapshots` only for isolated debugging, not as a normal completion path.

## Key behavior details

- **Email templates:** hidden config sheets/cells `D5` TO, `D7` CC, `D9` SUBJECT, `D11` BODY; `{{REPORT}}` marker splits intro/closing; `;` separates addresses; Alt+Enter for line breaks in BODY. Manager and director/dashboard subjects append current date `(m/d/yyyy)` if not already present (per current implementation).
- **Email copy ranges:** manager copies data range (header row down, bounded to data columns, excluding action buttons); dashboard copies used/chart extent (excluding email button shape). Outlook drafts created via COM/Word editor; wait/retry for editor readiness but within overall 30s budget.
- **Logging:** VBA `modLogging` (`LogMsg`/`LogErr`) and audit events should be written so per-run traces are diagnosable; Python runner logs outcome, marker, and per-run VBA trace when available.

## Environment notes (Windows/PowerShell)

- Working shell: PowerShell. Use `Get-ChildItem`/`Select-String` instead of `ls -la`/`grep`/`rg` if unavailable. Paths may contain spaces — quote them.
- Before COM/sync/snapshot operations, ensure no stale Excel instance blocks files; if Excel has template/report open, visible/property errors or hangs can occur — close/kill Excel and retry within 30s policy.
- Never commit secrets. Do not commit/push unless explicitly requested.
