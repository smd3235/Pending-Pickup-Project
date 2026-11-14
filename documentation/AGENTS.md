# Agent Rules — Pending Pickup Report Builder

These rules must be followed for every code change, no exceptions.

## 1. Git Workflow

- **Always commit and push to git after making changes.**
- **Always update `documentation/changelog.md` with new entries.**
- **Changelog entries must be newest to oldest.**
- **All Markdown files, including agent instructions, must be stored in `documentation/`.** Do not create Markdown files in the project root.

### Commit message format

- Use clear, descriptive commit messages that explain *what* changed and *why*.
- Reference the affected files or modules (e.g., `Pending_Pickup_Macro.bas`, `scripts/run-report.ps1`, `documentation/AGENTS.md`).
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

- **Run `scripts\test-compile.ps1` after every VBA code change to verify compilation.**
- **Run `scripts\sync_template.ps1` after compilation succeeds to update the template.**
- **Run `scripts\run-report.ps1` to test full report generation.**
- **Check `macro-debug.log` for any errors after a run.**
- **Check `log/answer-dialogs.log` for dialog handling issues.**
- **Kill any stray Excel processes after testing.**

### Debugging checklist

1. Compile check (`scripts\test-compile.ps1`).
2. Sync to template (`scripts\sync_template.ps1`).
3. Run report (`scripts\run-report.ps1`).
4. Check `macro-debug.log` for errors.
5. Check `Reports/` for output file.
6. Kill stray Excel processes.
7. Commit and push.

## 5. Cross-Project Audit (`audit-peer`)

- **Run `scripts/audit-peer.ps1` to compare this project with the sibling PM Project.** Use `-Sections scripts,documentation,commands,vba,rules` to narrow the audit, `-PeerPath <path>` when the peer is not discovered automatically, and `-Report documentation/cross-audit.md` to save a dated Markdown report.
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
- `scripts\sync_template.ps1` — Re-imports .bas into template.
- `scripts\test-compile.ps1` — Compile check.
- `scripts\run-report.ps1` — Full report run harness.
- `scripts\answer-dialogs.ps1` — Dialog auto-answerer.
- `scripts\audit-peer.ps1` — Read-only cross-project audit of portable files, VBA procedures, commands, and agent rules.
- `documentation/view-logs.ps1` — Reorders every runtime `.log` and `.txt` file newest first with dated headers.

## 7. Version Control

- **Never commit without explicit user request** unless following the mandatory commit rule above.
- **Always stage only intended files** — never `git add -A` unless explicitly asked.
- **Never commit secrets or keys to the repository.**
- **Never use `--amend` on a commit that has already been pushed.**
- **Never use force-push unless explicitly requested.**

## 7. Error Handling Patterns

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

## 8. Logging Standards

- All log entries must use full local date and time in `yyyy-MM-dd HH:mm:ss` format.
- All runtime log review output must be newest to oldest; run `documentation/view-logs.ps1` after log-producing work to reorder every runtime `.log` and `.txt` file in place.
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

## 9. Harness and Testing Rules

- **Excel must be killed within 10 seconds of a stall** — no exceptions.
- **All modal dialogs must be auto-answered** (CURRENT picker, PRIOR picker, error dialogs).
- **Error dialogs must be logged with full text** before being dismissed.
- **All test runs must verify output file creation** in `Reports/`.
- **All test runs must clean up temp files** (`_run_temp.xlsm`, `~$_run_temp.xlsm`).
- **All test runs must kill stray Excel processes** before exiting.

## 10. Deployment

- **Single-file distribution** — end users receive only the `.xlsm` template.
- **Never expose `.bas`, `scripts\sync_template.ps1`, `documentation/README.md`, `documentation/changelog.md`, or fake data to end users.**
- **Template must be self-contained** — logo embedded, module imported, button installed.
- **Registry settings must be pre-configured** — `CurrentFolder`, `PriorFolder`, `OutputFolder`.