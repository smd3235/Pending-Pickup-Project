# Agent Rules — Pending Pickup Report Builder

These rules must be followed for every code change, no exceptions.

## 1. Git Workflow

- **Python-first:** all file/script operations use Python (`python scripts\*.py`), never PowerShell. Do not use `powershell` for temp file creation/deletion, directory management, or any file manipulation — use Python scripts instead.
- **Always commit and push to git after making changes.**
- **Run full pipeline (`python scripts\pipeline.py`) before committing** — all quality-gate steps must pass (compile→sync→run→screenshots→scan/validate). Advisory steps (`pip check`, pytest coverage) report information but never fail the pipeline.
- **Always update `documentation/changelog.md` with new entries.**
- **Changelog entries must be newest to oldest.**
- **All Markdown files, including agent instructions, must be stored in `documentation/`.** Do not create Markdown files in the project root.

### Commit message format

- Use clear, descriptive commit messages that explain *what* changed and *why*.
- Reference the affected files or modules (e.g., `Pending_Pickup_Macro.bas`, `scripts/run-report.py`, `documentation/AGENTS.md`).
- Do not commit generated report files (`Reports/*.xlsx`) unless explicitly asked — they are build artifacts.

### Changelog format

Each entry must follow this format:

```
- **HH:MM:SS** — <what changed> and <why>.
  - <additional detail if needed>
```

- Timestamp is local time in `HH:MM:SS` format.
- Newest entries go at the top of the current date section.
- Group by date with `## YYYY-MM-DD` headers.
- Include file names, function names, and specific changes.

### Example

```markdown
## 2026-11-14

- **05:25:00** — Fixed "Type mismatch" runtime error in `ApplyAgeWarningFont` caused by `CDbl(v)` on formula
  strings (`=IF(...)`) in the Pickup Age column. Wrapped with `On Error Resume Next`.
- **05:25:01** — Added VBA-level logging to `macro-debug.log` with timestamps and procedure names for all
  key steps: `CreatePendingPickupReport`, `FormatManagerSheet`, `BuildDashboard`, etc.
```

## 2. Code Quality

- **Never introduce code that exposes or logs secrets and keys.**
- **Never commit secrets or keys to the repository.**
- **Follow security best practices.**
- **Use error handling (`On Error Resume Next` / `On Error GoTo 0`) for all risky operations.**
- **Log errors to `macro-debug.log` with procedure name, error description, error number, and line number.**
- **Bounds-check array access (`srcRow >= 1 And srcRow <= UBound(data, 1)`) before accessing.**
- **Check for `Nothing` before using object references.**
- **Use `IsNumeric()` before `CDbl()` on cell values that may contain formulas or text.**

## 3. VBA-Specific Rules

- **All procedures must have error handling** — either `On Error GoTo <label>` with a cleanup block, or `On Error Resume Next` for non-critical operations.
- **All functions that iterate over collections or arrays must bounds-check row indices.**
- **All `CDbl()`, `CLng()`, `CInt()` calls on cell values must be preceded by `IsNumeric()` or wrapped in error handling.**
- **All `Find` operations on worksheets must have `On Error Resume Next` fallbacks.**
- **All file I/O (SaveSetting, GetSetting, Dir, FileDateTime) must have error handling.**
- **All chart and formatting operations must have `On Error Resume Next` fallbacks.**

## 4. Testing and Debugging

- **Run `scripts\test-compile.py` after every VBA code change to verify compilation.**
- **Run `scripts\sync_template.py` after compilation succeeds to update the template.**
- **Run `scripts\run-report.py` to test full report generation.**
- **Run `python -m pytest tests/ -q` (no Excel launched) to verify pure-Python helpers.** Pytest coverage runs as an advisory pipeline step (`coverage report`) — it reports information but never fails the pipeline, because coverage's C tracer can crash on Windows ctypes Win32 callbacks in `vba_ocr_helpers.py`.
- **Run `scripts\vba-analyze.py` to statically scan the template for macros, stomping, XLM, and suspicious keywords.**
- **Check `macro-debug.log` for any errors after a run.**
- **Check `log/answer-dialogs.log` for dialog handling issues.**
- **Kill any stray Excel processes after testing.**

### Debugging checklist

1. Compile check (`scripts\test-compile.py`).
2. Sync to template (`scripts\sync_template.py`).
3. Run report (`scripts\run-report.py`).
4. Check `macro-debug.log` for errors.
5. Check `Reports/` for output file.
6. Kill stray Excel processes.
7. Commit and push.

## 5. Cross-Project Audit (`audit-peer`)

- **Run `scripts/audit-peer.py` to compare this project with the sibling PM Project.** Use `-Sections scripts,documentation,commands,vba,rules` to narrow the audit, `-PeerPath <path>` when the peer is not discovered automatically, and `-Report documentation/cross-audit.md` to save a dated Markdown report.
- **The audit is strictly read-only on both projects.** It must not copy, modify, format, or synchronize files. Porting is a separate, explicitly approved task.
- **Review both directions:** identify useful peer-only files and procedures, plus project-only material the peer may want. Do not recommend copying code that duplicates existing functionality or depends on the other project's data model without adaptation.
- **Agent-rule changes must be reviewed separately.** Compare `documentation/AGENTS.md`, changelog conventions, testing rules, and deployment constraints; update this file only when a rule is actually adopted by this project.
- **Any approved port must be adapted to this project's paths and data sources, then pass the full validation loop:** compile check, template synchronization, end-to-end report run, log review, changelog update, commit, and push.
- **Never copy secrets, generated reports, runtime logs, or machine-specific temp paths during an audit.**

## 6. File Organization

- `Pending_Pickup_Macro.bas` — Main VBA module (source of truth).
- `Pending Pickup Template.xlsm` — Template with embedded module and START HERE button.
- `Reports/` — Generated report output (build artifacts, not committed).
- `log/` — Runtime logs (macro-debug.log, answer-dialogs.log, etc.).
- `documentation/` — Markdown documentation and image assets, including `README.md`, `changelog.md`, and logo source files.
- `documentation/AGENTS.md` — Agent rules.
- `documentation/changelog.md` — Change history (newest to oldest).
- `documentation/README.md` — User documentation.
- `documentation/aoti-logo-2023.svg` — Vector logo source.
- `documentation/aoti-logo-2023.png` — Raster logo source.
- `scripts\sync_template.py` — Re-imports .bas into template.
- `scripts\test-compile.py` — Compile check.
- `scripts\run-report.py` — Full report run harness.
- `scripts\answer-dialogs.py` — Dialog auto-answerer. When it hits any unrecognized VBA error dialog (compile or runtime), launches `vba-error-ocr.py ocr-live --pid <pid>` to snapshot the live VBA UI before dismissing, and logs the parsed error number.
- `scripts\audit-peer.py` — Read-only cross-project audit of portable files, VBA procedures, commands, and agent rules.
- `scripts\vba-error-ocr.py` — Captures VBA error UI and OCRs the highlighted code band, snapping OCR text to `vba/*.bas` source lines. Subcommands: `collect` (launch probe, run macro, capture on first dialog, kill Excel to unblock), `ocr-live --pid <n>` (attach to running Excel during a real run; captures dialog + VBE + code pane + desktop + all VBA windows without killing Excel), `ocr-all` (batch summary of all captures under `log\ocr\`), `analyze` (OCR one screenshot). VBA's compile-error dialog ignores window-message clicks, so `collect` kills Excel as soon as the dialog appears to unblock the blocked macro call. Artifacts under `log\ocr\`, summary in `log\vba-error-ocr.log`.
- `scripts\vba-analyze.py` — oletools-based static VBA analysis of `Pending Pickup Template.xlsm`: reports macro presence, VBA stomping (source/p-code mismatch), XLM (Excel 4.0) macros, and suspicious keyword hits (Shell/CreateObject/Environ/etc.). Run as part of pipeline quality gates.
- `scripts\vba_ocr_helpers.py` — dependency-light shared helpers used by `vba-error-ocr.py` and the test suite (normalize, snap_to_source, detect_bands, flatten_selection, preprocess_variants, dialog_child_texts). Lazy PIL import keeps import fast for testing.
- `scripts\screenshot-diff.py` — Perceptual-hash (imagehash) screenshot regression check for `documentation/screenshots/`; compares against `documentation/screenshots/_baseline/` and fails the pipeline on regressions (>5 hamming distance). `--refresh-baseline` re-establishes the baseline from current screenshots.
- `scripts\ocr-winrt.ps1` — Windows.Media.Ocr WinRT bridge used by `scripts\vba-error-ocr.py` (OCR engine only; all file/script work stays in Python).
- `scripts\mypy.ini` — mypy config; pipeline runs `mypy --follow-imports=skip` on owned files.
- `requirements.txt` — Pinned direct dependencies (Pillow, pywin32, types-psutil, types-pywin32, oletools, pytest, ruff, black, mypy, ImageHash, GitPython, openpyxl, psutil, rich, coverage).
- `tests\` — pytest suite (no Excel launched): `tests/conftest.py`, `tests/test_helpers.py` (normalize/snap/detect_bands/flatten_selection/dialog_child_texts), `tests/test_cli.py` (CLI smoke tests). Runs in the pipeline as `unit tests (pytest)`.
- `scripts\temp\` — Temporary scratch directory for agent inspection/debug scripts. **Never use shared temp folders** (e.g., `C:\Users\*\AppData\Local\Temp\kilo`) — other projects share them and will confuse or delete files. Create and clean up temp files inside `scripts\temp\`. Never leave temp scripts, logs, or extracted artifacts in shared temp folders — they will confuse other projects.
- `documentation\view-logs.py` — Reorders every runtime `.log` and `.txt` file newest first with dated headers.

## 7. Version Control

- **Never commit without explicit user request** unless following the mandatory commit rule above.
- **Always stage only intended files** — never `git add -A` unless explicitly asked.
- **Never commit secrets or keys to the repository.**
- **Never use `--amend` on a commit that has already been pushed.**
- **Never use force-push unless explicitly requested.**

## 8. Error Handling Patterns

### For procedures that should not fail silently

```vba
Private Sub MyProcedure()
    On Error GoTo CleanFail
    ' ... code ...
    Exit Sub
CleanFail:
    LogErr "MyProcedure"
    ' ... cleanup ...
End Sub
```

### For procedures where failure is non-critical

```vba
Private Sub MyNonCriticalProcedure()
    On Error Resume Next
    ' ... risky code ...
    On Error GoTo 0
End Sub
```

### For array access with bounds checking

```vba
If srcRow >= 1 And srcRow <= UBound(data, 1) Then
    ' safe to access data(srcRow, col)
End If
```

### For cell value conversion

```vba
If IsNumeric(v) And Len(Trim$(CStr(v))) > 0 Then
    If CDbl(v) > 30 Then ' ... '
End If
```

## 9. Logging Standards

- All log entries must use full local date and time in `yyyy-MM-dd HH:mm:ss` format.
- All runtime log review output must be newest to oldest; run `documentation/view-logs.py` after log-producing work to reorder every runtime `.log` and `.txt` file in place.
- All errors must be logged to `macro-debug.log` with:
  - Timestamp (`Format(Now, "yyyy-mm-dd hh:n:ss")`)
  - Procedure name
  - Error description (`Err.Description`)
  - Error number (`Err.Number`)
  - Line number (`Erl`) when available
- Key milestones must be logged:
  - `CreatePendingPickupReport started`
  - `CURRENT selected: <path>`
  - `PRIOR picker (suggested: <path>)`
  - `priorPath=False? <bool>`
  - `priorCount=<n> hasPrior=<bool>`
  - `Aging: 30=<n> 60=<n> 90=<n> missing=<n>`
  - `Output workbook created`
  - `Creating manager tab: <name>`
  - `FormatManagerSheet start: <name> pickupCol=<n>`
  - `FormatManagerSheet: <name> done stats`
  - `  lastDataRow=<n>`
  - `  ApplyDataFormatting`
  - `  ApplyDateFormatsToOutput`
  - `  ApplyAgeWarningFont`
  - `  ApplyRepBanding`
  - `  HighlightMissingPickupDates`
  - `  FormatYesNoColumns`
  - `  formatting done`
  - `  ApplyDateFormatsToOutput: <name>`
  - `  MatchSourceColumnWidthsForManager: <name>`
  - `  manager tab done: <name>`
  - `All manager tabs done`
  - `BuildDashboard start`
  - `BuildDashboard done`
  - `Raw Data sheet created`
  - `Saving to: <path>`
  - `Save done`
  - `FAILED: <error> (Err=<n>)`
  - `=== CleanFail exit ===`
  - `=== CreatePendingPickupReport started ===`
  - `EnsureMacroEnabledWorkbook OK`

## 10. Harness and Testing Rules

- **Excel must be killed within 10 seconds of a stall** — no exceptions.
- **All modal dialogs must be auto-answered** (CURRENT picker, PRIOR picker, error dialogs).
- **Error dialogs must be logged with full text** before being dismissed.
- **All test runs must verify output file creation** in `Reports/`.
- **All test runs must clean up temp files** (`_run_temp.xlsm`, `~$_run_temp.xlsm`).
- **All test runs must kill stray Excel processes** before exiting.
- **Screenshots are regenerated on every test run:** before running `scripts\snapshot-report.py`, delete all existing PNGs in `documentation/screenshots/` so the committed screenshots always reflect the latest report output, never a stale mix of old and new sheets.
- **Verify screenshot count:** after regeneration, confirm all expected screenshots are present (14 report sheets + `template-start-here.png` + `template-manager-email-template.png` + `template-director-email-template.png` = 17 total).
- **Temp file cleanup:** after any debugging or inspection, delete all files created in `scripts\temp\`. Never leave temp scripts, logs, or extracted artifacts in shared temp folders — they will confuse other projects.

## 11. Deployment

- **Single-file distribution** — end users receive only the `.xlsm` template.
- **Never expose `.bas`, `scripts\sync_template.py`, `documentation/README.md`, `documentation/changelog.md`, or fake data to end users.**
- **Template must be self-contained** — logo embedded, module imported, button installed.
- **Registry settings must be pre-configured** — `CurrentFolder`, `PriorFolder`, `OutputFolder`.

## 12. Pipeline Quality Gates & Evaluated Tools

### Pipeline quality gates (run in order by `scripts\pipeline.py`)
1. **deps (pip check)** — advisory only; reports shared-environment dependency conflicts (opentelemetry, safety, etc.) but never fails the pipeline.
2. **lint (ruff)** — `ruff check --select E9,F821,F722 scripts/ tests/`; fails the pipeline on real bug categories (syntax, undefined names, forward-annotation errors).
3. **unit tests (pytest)** — `pytest tests/ -q`; fails the pipeline on test failure. No Excel is launched.
4. **coverage report** — advisory only; `coverage run -m pytest tests/` + `coverage report`. Never fails the pipeline — coverage's C tracer can crash on Windows ctypes Win32 callbacks (`vba_ocr_helpers.py:dialog_child_texts`), and that is expected.
5. **vba static analysis (vba-analyze.py)** — oletools scan of the template for macros, stomping, XLM, suspicious keywords.
6. **types (mypy)** — `mypy --follow-imports=skip` on owned scripts + tests.
7. **compile (test-compile.py)** — injects a probe macro, runs it (forces full-project compile), captures any modal error dialog via PrintWindow, OCRs the highlighted code line, snaps to `vba/*.bas` sources.
8. **sync (sync_template.py)** — re-imports .bas modules into the template.
9. **run-report (run-report.py)** — full report generation, no screenshots.
10. **snapshot-report** — regenerates all QA screenshots (required; 120s cap).
11. **screenshot-regression (screenshot-diff.py)** — perceptual-hash compare of regenerated screenshots against `documentation/screenshots/_baseline/`; fails the pipeline on regressions (>5 hamming distance).
12. **monitor-logs / analyze-trace / validate-report / check-logs** — fast log/report scans.
13. **audit-code** — VBA code audit.

### Tools evaluated but NOT integrated as pipeline gates
- **bandit** — runs, but 209 Low-severity findings are inherent to this Windows COM automation project (legitimate `try/except: pass` for COM error handling, `subprocess` for PowerShell OCR/taskkill). Would fail the pipeline on noise, not real problems.
- **safety** — crashes (`TypeError: 'Timeout' object`); requires network/auth to the safety DB, unavailable here.
- **coverage** — used as an **advisory** step only (see gate #4). Its C tracer access-violates on Windows ctypes Win32 callbacks, so it cannot be a hard gate.
- **vulture** — likely false-positives on dynamically-called VBA/COM code; not worth a hard gate.
- **detect-secrets** — module not importable despite being listed as installed.
- **black** — installed and available (`python -m black --check scripts/ tests/`), but not a hard gate because it would reformat a large amount of existing code. Run manually when desired.