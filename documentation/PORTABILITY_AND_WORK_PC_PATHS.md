# Portability & Work-PC Path Issues — Review and Fix Guide

This document covers the "embedded file locations" issue that can prevent the macro from running correctly on a different PC (e.g. work machine). When paths are hardcoded to a developer machine, the macro may try to read/write in locations that do not exist, point to the wrong folders, or fail silently under `On Error Resume Next`.

**Goal:** the template/macro must run on any machine with only `Pending Pickup Template.xlsm` (or a locally rebuilt `.xlsm`) and user-selected source files, with no dependency on developer-specific paths.

---

## 1. Symptoms on another PC

- Macro errors, hangs, or "nothing happens" when logging or auto-run files reference a missing drive/folder.
- Log/completion marker files not created where expected (or attempts to write to `C:\Users\<other-user>\...` dev paths fail).
- Prompts/default folders open to non-existent locations (if stale remembered paths are treated as valid).
- Output not saved next to the workbook as intended if path logic is wrong (current output path is relative; see below).

---

## 2. Embedded / hardcoded locations found (current state)

### A. VBA constants with absolute developer paths (**must fix**)

| File | Constant | Current value (dev machine) | Used for |
|---|---|---|---|
| `modDeclarations.bas` | `LOG_PATH` | `C:\Users\smd32\Downloads\Pending Pickup Project\log\macro-debug.log` | Main debug/error log (`LogMsg`, `LogErr`, `modAudit.bas` audit append) |
| `modDeclarations.bas` | `COMPLETION_MARKER_PATH` | `...\log\pending-pickup-complete.txt` | Completion marker written by `WriteCompletionMarker` |
| `modDeclarations.bas` | `AUTO_SOURCE_PATH_FILE` | `...\log\auto-source.txt` | Unattended auto-source input (dev/test harness) |
| `modDeclarations.bas` | `AUTO_SKIP_PRIOR_FILE` | `...\log\auto-skip-prior.txt` | Unattended skip-prior flag (dev/test harness) |
| `Pending_Pickup_Macro.bas` | same names | same absolute paths | Legacy monolithic copy (keep consistent if retained) |

**Impact:** `modLogging.bas` and `modAudit.bas` derive directories via `Left$(..., InStrRev(...,"\"))` and call `MkDir`/`Open` on these paths. On a work PC the parent path likely does not exist; with `On Error Resume Next` logging can be skipped or behave unpredictably. `BuildRunLogPath` also depends on `LOG_PATH`.

### B. Relative / workbook-based paths (**portable, OK**)

| Location | Behavior | Status |
|---|---|---|
| `modDialog.bas` → `BuildOutputPath` | `reportsDir = masterWb.Path & "\Reports"`; creates if missing; saves report there | OK if workbook is in a writable local folder |
| `modDialog.bas` → `SelectExcelFile` | If `suggestedFolder` empty/invalid, falls back to `ThisWorkbook.Path`; validates with `Dir$(..., vbDirectory)` | Generally OK; depends on saved folder validity (see C) |

### C. Registry-persisted folders (machine/user-specific, can be stale)

| Key | Registry path | Notes |
|---|---|---|
| `CurrentFolder` | `HKCU\Software\VB and VBA Program Settings\PendingPickupReport\Folders\CurrentFolder` | Remembers last source folder |
| `PriorFolder` | same hive, `PriorFolder` | Remembers prior report folder; also used by `FindMostRecentPendingPickupReport` suggestion |
| `OutputFolder` | documented in README as `...\PendingPickupReport\Settings\OutputFolder` (verify actual key used in code) | If pointed at a dev path, save/output suggestions can be wrong on another PC |

**Impact:** not "embedded in code", but can embed a bad location per-user. `SelectExcelFile` should ignore non-existent suggested folders and fall back to `ThisWorkbook.Path` (current code does check existence before applying in some cases—verify all call sites). Ensure invalid/stale paths never block the run.

### D. Documentation/scripts with example absolute paths (**non-runtime but can confuse**)

- `README.md`, docs, and helper scripts may show `C:\Users\smd32\...` examples. These do not affect macro execution on a work PC but should be treated as examples only.

### E. Template/workbook-embedded content to verify

- Shapes/buttons: must call macro names (`CreatePendingPickupReport`, `SetupEmailCC`, etc.), not external file paths — OK if assigned by name.
- Hyperlinks, queries, external data connections, or defined names pointing to local files: **audit template** for any machine-specific links (none expected, but verify before distribution).
- START HERE text: avoid hardcoding dev paths; keep instructions generic.

---

## 3. Recommended fixes (implementation)

### 1) Remove absolute log/marker/auto paths — derive at runtime
Replace path constants with functions that resolve to a writable, project-local location:

- **Preferred:** `<ThisWorkbook.Path>\log\...` (keeps artifacts with the workbook/project; requires workbook saved in a writable local folder—already recommended for work PCs).
- **Fallback:** if `ThisWorkbook.Path` is empty/unwritable (e.g. unsaved/temp edge cases), use a user-writable location such as `Environ$("LOCALAPPDATA") & "\PendingPickupReport\log\"` or `Environ$("TEMP") & "\PendingPickupReport\log\"`.

Suggested helpers (e.g. in `modLogging.bas` or `modUtils.bas`):
- `GetLogDirectory() As String` — returns ensured directory path.
- `GetLogPath()`, `GetCompletionMarkerPath()`, `GetAutoSourcePath()`, `GetAutoSkipPriorPath()` — build from `GetLogDirectory()`.
- `EnsureFolderExists(ByVal folderPath As String)` — safe create (handle trailing slash, existing dir, errors).

Update all references: `modDeclarations.bas` (remove/replace constants), `modLogging.bas`, `modAudit.bas` (`Open LOG_PATH`), `modCore.bas` (auto-file checks and marker messages). Keep behavior the same, only path source changes.

### 2) Make logging resilient
- Never let logging failure break report generation: keep `On Error Resume Next` around file I/O but ensure directory resolution cannot produce an empty/invalid rooted path without handling.
- If the log directory cannot be created, skip file logging gracefully (optionally `Debug.Print`) rather than raising user-visible errors during normal runs.

### 3) Harden file/folder selection and saved folders
- In `SelectExcelFile`, if `suggestedFolder` or `suggestedFile` path does not exist, ignore it and default to `ThisWorkbook.Path` (then user profile/Documents as last resort if needed).
- Optionally detect obviously invalid/stale saved paths (empty, drive missing, `Dir` fails) and do not pre-fill them.
- Keep `BuildOutputPath` relative to the active/master workbook path (current behavior) and ensure `Reports` creation handles failure with a clear message if the workbook folder is read-only.

### 4) Treat auto-run files as dev/test harness only
- `auto-source.txt` / `auto-skip-prior.txt` are for automated runs; once moved under the derived log directory they remain portable. Do not require them for normal interactive use.
- Ensure missing auto files simply mean "not in auto mode" (current intent).

### 5) Legacy module parity
- `Pending_Pickup_Macro.bas` contains the same hardcoded paths; if it is kept as a reference/build source, update it to match the portable implementation to avoid reintroducing embedded paths.

### 6) Template/build/package hygiene
- After VBA changes: `python scripts/sync_template.py` (rebuilds modules, `Template (No Macros).xlsx`, `Project-Local-Build.zip`) and `python scripts/test-compile.py`.
- Verify `START HERE` button assignments and absence of external links/paths in the template.
- Distribute/run from a **local writable folder** (not blocked network/temp download location) per org policy.

---

## 4. Verification checklist (on a non-dev/work PC)

- [ ] Copy/build `.xlsm` to a local writable folder (not the dev path) and open.
- [ ] Confirm no references to `C:\Users\smd32\...` in VBA modules/constants (search project).
- [ ] Run report interactively: select current export, skip/select prior, save completes to `<workbook folder>\Reports\...`.
- [ ] Confirm logs/marker are created under the resolved portable log directory (e.g. `<workbook folder>\log\`) and no errors due to missing dev folders.
- [ ] Clear/point registry folder settings to invalid paths and confirm picker still opens a valid default folder.
- [ ] Run `scripts/test-compile.py` on dev after changes; smoke-test generated output.

---

## 5. Related files

- `modDeclarations.bas` — hardcoded path constants (primary source of issue)
- `modLogging.bas` — `LogMsg`, `LogErr`, `BuildRunLogPath`, `WriteCompletionMarker`
- `modAudit.bas` — appends to `LOG_PATH`
- `modCore.bas` — reads auto files, writes marker/log messages, output/save flow
- `modDialog.bas` — `SelectExcelFile`, folder registry settings, `BuildOutputPath`, prior-report discovery
- `Pending_Pickup_Macro.bas` — legacy monolithic copy with same hardcoded paths
- `README.md`, `documentation/README.md` — distribution/work-PC guidance
- `scripts/sync_template.py`, `scripts/test-compile.py` — rebuild/validation

## 6. Notes

- `BuildOutputPath` is already relative to `masterWb.Path` and is the correct pattern to follow for other runtime artifacts.
- Registry values are per-user; "wrong spot" reports from another PC can come from stale saved folders even when code is portable—handle invalid paths defensively.
- This is separate from macro-block/Mark-of-the-Web issues; blocked downloads require the macro-free local build package (`Project-Local-Build.zip`) as documented in project README/docs.
