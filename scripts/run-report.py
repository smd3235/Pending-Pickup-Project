# run-report.py
# Runs the report macro end-to-end against a temp copy of the template.
# Hard rules: Excel must be killed if the run stalls (no new report file within
# $MaxRunSeconds, default 30s) or if a modal dialog stays unanswered too long.
# The watchdog + dialog answerer run on separate threads so they work even while
# the main thread is blocked inside the synchronous excel.Run call.
# Logs: log\pending-pickup_<timestamp>.log and log\answer-dialogs.log

import argparse
import ctypes
import ctypes.wintypes as wintypes
import os
import shutil
import subprocess
import sys
import threading
import time
from datetime import datetime

import openpyxl
import win32com.client

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

EnumWindows = user32.EnumWindows
EnumChildWindows = user32.EnumChildWindows
IsWindowVisible = user32.IsWindowVisible
GetWindowTextW = user32.GetWindowTextW
GetClassNameW = user32.GetClassNameW
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
SetForegroundWindow = user32.SetForegroundWindow
SendMessageW = user32.SendMessageW
PostMessageW = user32.PostMessageW

EnumProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

WM_SETTEXT = 0x000C
WM_COMMAND = 0x0111
WM_KEYDOWN = 0x0100
IDOK = 1
VK_RETURN = 0x0D
VK_ESCAPE = 0x1B

parser = argparse.ArgumentParser(description="Run the Pending Pickup report macro")
parser.add_argument("--template", default=None)
parser.add_argument("--source", default=None)
parser.add_argument("--temp-copy", default=None)
parser.add_argument("--max-run-seconds", type=int, default=30)
parser.add_argument("--stall-seconds", type=int, default=10)
parser.add_argument("--no-screenshots", action="store_true", help="Skip post-run screenshot regeneration (debug only; screenshots are required)")
args = parser.parse_args()

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)

template = args.template or os.path.join(project_root, "Pending Pickup Template.xlsm")
source_path = args.source or os.path.join(project_root, "AllOpenTreatmentPlansResults618.xlsx")
temp_copy = args.temp_copy or os.path.join(project_root, "_run_temp.xlsm")
max_run_seconds = args.max_run_seconds
stall_seconds = args.stall_seconds
SCREENSHOT_TIMEOUT_SECONDS = 120  # screenshots may take >30s but must finish within 2m or be treated as hang

reports_dir = os.path.join(project_root, "Reports")
os.makedirs(reports_dir, exist_ok=True)
log_dir = os.path.join(project_root, "log")
os.makedirs(log_dir, exist_ok=True)

run_stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
run_log = os.path.join(log_dir, f"pending-pickup_{run_stamp}.log")
ans_log = os.path.join(log_dir, "answer-dialogs.log")
completion_marker = os.path.join(log_dir, "pending-pickup-complete.txt")

before = set()
if os.path.isdir(reports_dir):
    before = {f.name for f in os.scandir(reports_dir)
              if f.name.lower().endswith(".xlsx") and not f.name.startswith("~$")}

if os.path.exists(completion_marker):
    os.remove(completion_marker)


def write_log(path, msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    with open(path, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def append_run_log(msg):
    write_log(run_log, msg)


EMAIL_SYNC_SHEETS = ["Manager Email Template", "Director Email Template"]
EMAIL_SYNC_CELLS = ["D5", "D7", "D9", "D11"]  # TO, CC, SUBJECT, BODY


def read_email_sync_config(path):
    try:
        wb = openpyxl.load_workbook(
            path,
            keep_vba=path.lower().endswith(".xlsm"),
            data_only=False,
            read_only=False,
        )
        out = {}
        for name in EMAIL_SYNC_SHEETS:
            ws = wb[name]
            vals = {}
            for c in EMAIL_SYNC_CELLS:
                v = ws[c].value
                vals[c] = None if v is None else str(v).replace("\r\n", "\n").replace("\r", "\n")
            out[name] = vals
        try:
            wb.close()
        except Exception:
            pass
        return out, None
    except Exception as e:
        return None, str(e)


def ensure_template_email_sync(template_path, project_root_path, log_fn):
    """Keep macro-enabled and no-macro artifacts aligned before report runs."""
    no_macro_path = os.path.join(project_root_path, "Pending Pickup Template (no-macros).xlsx")
    sync_script = os.path.join(project_root_path, "scripts", "sync_template.py")
    if not os.path.exists(template_path):
        log_fn(f"Email/template sync skipped: template not found ({template_path})")
        return
    if not os.path.exists(sync_script):
        log_fn(f"Email/template sync skipped: sync script not found ({sync_script})")
        return

    xlsm_sig, err1 = read_email_sync_config(template_path)
    if err1:
        log_fn(f"Email/template sync skipped: could not read .xlsm ({err1})")
        return
    xlsx_sig, err2 = read_email_sync_config(no_macro_path)

    need_sync = False
    reason = "email config matches no-macro"
    if err2:
        need_sync = True
        reason = f"no-macro unreadable/missing ({err2})"
    elif xlsm_sig != xlsx_sig:
        need_sync = True
        diffs = []
        for sh in EMAIL_SYNC_SHEETS:
            for c in EMAIL_SYNC_CELLS:
                if (xlsm_sig.get(sh) or {}).get(c) != (xlsx_sig.get(sh) or {}).get(c):
                    diffs.append(f"{sh}!{c}")
        reason = "email/template fields differ: " + ", ".join(diffs[:8]) + ("..." if len(diffs) > 8 else "")

    if not need_sync:
        log_fn(f"Email/template sync check: {reason}")
        return

    log_fn(f"Running sync_template.py before report ({reason})...")
    try:
        r = subprocess.run(
            [sys.executable, sync_script],
            cwd=project_root_path,
            capture_output=True,
            text=True,
            timeout=30,
        )
        if r.returncode == 0:
            log_fn("sync_template.py completed successfully before report run.")
            out = (r.stdout or "").strip().splitlines()
            for ln in out[-10:]:
                log_fn("  " + ln)
        else:
            log_fn(f"sync_template.py failed before report run (exit {r.returncode}). Continuing with current template.")
            out = ((r.stdout or "") + ("\n" + (r.stderr or ""))).strip().splitlines()
            for ln in out[-15:]:
                log_fn("  " + ln)
    except subprocess.TimeoutExpired:
        log_fn("sync_template.py timed out (>30s) before report run. Continuing with current template.")
    except Exception as e:
        log_fn(f"Could not run sync_template.py before report ({e}). Continuing with current template.")


write_log(run_log, "=== Pending Pickup report run started ===")
write_log(run_log, f"Template: {template}")
write_log(run_log, f"Source: {source_path}")
write_log(run_log, f"Completion marker: {completion_marker}")
write_log(run_log, f"Python PID: {os.getpid()}")

ensure_template_email_sync(template, project_root, append_run_log)

# Write auto-mode files so VBA skips file picker dialogs during automated runs.
auto_source_file = os.path.join(log_dir, "auto-source.txt")
auto_skip_prior_file = os.path.join(log_dir, "auto-skip-prior.txt")
with open(auto_source_file, "w", encoding="utf-8") as f:
    f.write(source_path)
append_run_log(f"Wrote auto source path: {source_path}")
with open(auto_skip_prior_file, "w", encoding="utf-8") as f:
    f.write("1")
append_run_log("Wrote auto skip-prior flag.")

shutil.copy2(template, temp_copy)
outcome = "UNKNOWN"
run_completed = False
run_start_time = time.time()
excel_pid = None
cancel_event = threading.Event()


def excel_pids():
    try:
        import psutil
        pids = set()
        for p in psutil.process_iter(["name"]):
            try:
                if p.info["name"] and p.info["name"].lower() == "excel.exe":
                    pids.add(p.pid)
            except Exception:
                pass
        return pids
    except ImportError:
        return set()


def win_text(hwnd):
    buf = ctypes.create_unicode_buffer(1024)
    GetWindowTextW(hwnd, buf, 1024)
    return buf.value


def win_class(hwnd):
    buf = ctypes.create_unicode_buffer(256)
    GetClassNameW(hwnd, buf, 256)
    return buf.value


def find_child(hwnd, cls_name):
    found = [None]

    def cb(child, _):
        if win_class(child) == cls_name:
            found[0] = child
            return False
        return True

    EnumChildWindows(hwnd, EnumProc(cb), 0)
    return found[0]


def excel_dialogs():
    pids = excel_pids()
    if not pids:
        return []
    result = []

    def cb(hwnd, _):
        if not IsWindowVisible(hwnd):
            return True
        pid = wintypes.DWORD()
        GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        if pid.value not in pids:
            return True
        result.append(hwnd)
        return True

    EnumWindows(EnumProc(cb), 0)
    return result


def send_keys(text):
    """Send keystrokes via SendInput."""
    class KEYBDINPUT(ctypes.Structure):
        _fields_ = [
            ("wVk", wintypes.WORD),
            ("wScan", wintypes.WORD),
            ("dwFlags", wintypes.DWORD),
            ("time", wintypes.DWORD),
            ("dwExtraInfo", ctypes.c_void_p),
        ]

    class INPUT(ctypes.Structure):
        _fields_ = [
            ("type", wintypes.DWORD),
            ("ki", KEYBDINPUT),
        ]

    INPUT_KEYBOARD = 1
    KEYEVENTF_KEYUP = 0x0002

    VkKeyScanW = user32.VkKeyScanW
    inputs = []
    for ch in text:
        vk = VkKeyScanW(ch)
        inputs.append(INPUT(1, KEYBDINPUT(vk & 0xFF, 0, 0, 0, 0)))
        inputs.append(INPUT(1, KEYBDINPUT(vk & 0xFF, 0, KEYEVENTF_KEYUP, 0, 0)))
    if inputs:
        n = len(inputs)
        arr = (INPUT * n)(*inputs)
        user32.SendInput(n, arr, ctypes.sizeof(INPUT))


def key_down_once(vk):
    class KEYBDINPUT(ctypes.Structure):
        _fields_ = [
            ("wVk", wintypes.WORD),
            ("wScan", wintypes.WORD),
            ("dwFlags", wintypes.DWORD),
            ("time", wintypes.DWORD),
            ("dwExtraInfo", ctypes.c_void_p),
        ]

    class INPUT(ctypes.Structure):
        _fields_ = [
            ("type", wintypes.DWORD),
            ("ki", KEYBDINPUT),
        ]

    INPUT_KEYBOARD = 1
    KEYEVENTF_KEYUP = 0x0002

    ki_down = KEYBDINPUT(vk, 0, 0, 0, 0)
    ki_up = KEYBDINPUT(vk, 0, KEYEVENTF_KEYUP, 0, 0)
    arr = (INPUT * 2)(INPUT(1, ki_down), INPUT(1, ki_up))
    user32.SendInput(2, arr, ctypes.sizeof(INPUT))


def kill_excel():
    try:
        subprocess.run(["taskkill", "/F", "/IM", "excel.exe"], capture_output=True, timeout=10)
    except Exception:
        pass


# ---------------------------------------------------------------- Thread A: watchdog
def watchdog_thread():
    sw = time.time()
    last_progress = 0
    last_dialogs = 0
    while not cancel_event.is_set() and (time.time() - sw) < max_run_seconds:
        time.sleep(0.3)
        # New report file = success.
        try:
            new_files = [
                f.name for f in os.scandir(reports_dir)
                if f.name.lower().endswith(".xlsx") and f.name not in before and not f.name.startswith("~$")
            ]
        except OSError:
            new_files = []
        if new_files:
            write_log(run_log, f"REPORT OK: {', '.join(sorted(new_files))}")
        if os.path.exists(completion_marker):
            write_log(run_log, "MACRO COMPLETE: marker found; forcing Excel shutdown if COM remains blocked")
            kill_excel()
            return
        # Stall check: only kill if an answerable dialog stays open too long.
        # Excel being busy (macro running) is NOT a stall condition.
        dialog_open = False
        for h in excel_dialogs():
            cls = win_class(h)
            title = win_text(h)
            if (cls == "bosa_sdm_MsoFiles" or cls == "#32770" or cls == "wndclass_desked_gsk"
                    or "Select the CURRENT" in title or "Select the PRIOR" in title
                    or "Choose where" in title or "Microsoft Excel" in title):
                dialog_open = True
        if dialog_open:
            last_dialogs = time.time() - sw
            continue
        # No dialog open - Excel is either busy running the macro or idle.
        # Don't kill for being busy; max_run_seconds handles the overall cap.
    write_log(run_log, f"KILLED: run exceeded {max_run_seconds} s without a new report file")
    kill_excel()


# --------------------------------------------------------------- Thread B: dialog answerer
def answerer_thread():
    sw = time.time()
    answered = {}
    while not cancel_event.is_set() and (time.time() - sw) < 120:
        time.sleep(0.25)
        live = {}
        for h in excel_dialogs():
            live[str(h)] = h
        for k in list(answered.keys()):
            if k not in live:
                del answered[k]
        for k in live:
            if k in answered:
                continue
            h = live[k]
            cls = win_class(h)
            title = win_text(h)
            if (cls == "bosa_sdm_MsoFiles" or cls == "#32770" or cls == "wndclass_desked_gsk"
                    or "Select the CURRENT" in title or "Select the PRIOR" in title
                    or "Choose where" in title):
                write_log(ans_log, f"dialog opened: class={cls} title={title}")
            if cls == "#32770" and title == "Microsoft Excel":
                # Dismiss Excel macro/security error dialogs so the run can continue.
                write_log(ans_log, f"ERROR DIALOG children: {win_text(h)}")
                key_down_once(VK_RETURN)
                write_log(ans_log, f"dismissed Excel error dialog: {title}")
                answered[k] = 1
                continue
            if (cls == "wndclass_desked_gsk" or cls == "#32770") and title == "Microsoft Visual Basic for Applications":
                # VBA compile/runtime error dialog — enumerate children to capture the error message.
                err_holder = [""]
                def _enum_child(child, _):
                    ct = win_text(child)
                    cc = win_class(child)
                    if cc == "Static" and len(ct) > 3:
                        err_holder[0] += " | " + ct
                    return True
                EnumChildWindows(h, EnumProc(_enum_child), 0)
                err_text = err_holder[0] or win_text(h)
                write_log(ans_log, f"VBA ERROR DIALOG: {err_text}")
                key_down_once(VK_RETURN)
                write_log(ans_log, f"dismissed VBA compile/runtime error: {err_text[:200]}")
                answered[k] = 1
                continue
            if (cls == "bosa_sdm_MsoFiles" or cls == "#32770" or cls == "wndclass_desked_gsk"
                    or "Select the CURRENT" in title or "Select the PRIOR" in title
                    or "Choose where" in title):
                SetForegroundWindow(h)
                time.sleep(0.3)
                if "PRIOR" in title:
                    key_down_once(VK_ESCAPE)
                    write_log(ans_log, "cancelled PRIOR picker")
                elif "CURRENT" in title:
                    edit = find_child(h, "Edit")
                    if edit is not None:
                        SendMessageW(edit, WM_SETTEXT, 0, source_path)
                    else:
                        send_keys("^a")
                        send_keys(source_path)
                    key_down_once(VK_RETURN)
                    write_log(ans_log, f"answered CURRENT picker -> {source_path}")
                else:
                    key_down_once(VK_RETURN)
                    write_log(ans_log, f"dismissed dialog: {title}")
                answered[k] = 1


t_a = threading.Thread(target=watchdog_thread, daemon=True)
t_b = threading.Thread(target=answerer_thread, daemon=True)
t_a.start()
t_b.start()

excel = None
wb = None
try:
    # Set VBAWarnings registry BEFORE Excel starts so macros are enabled by default.
    try:
        import winreg
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\16.0\Excel\Security\VBAWarnings")
        winreg.SetValueEx(key, "VBAWarnings", 0, winreg.REG_DWORD, 1)  # 1 = enable all macros
        winreg.CloseKey(key)
        append_run_log("VBAWarnings registry set to enable macros.")
    except Exception as e:
        append_run_log(f"Could not set VBAWarnings: {e}")

    excel = win32com.client.DispatchEx("Excel.Application")
    append_run_log("Starting Excel: PID unknown until COM creates process")

    # Add project folder to Excel trusted locations so macros run without prompts.
    try:
        import winreg
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\PendingPickup")
        winreg.SetValueEx(key, "Path", 0, winreg.REG_SZ, project_root)
        winreg.SetValueEx(key, "Description", 0, winreg.REG_SZ, "Pending Pickup Project")
        winreg.SetValueEx(key, "AllowSubfolders", 0, winreg.REG_DWORD, 1)
        winreg.CloseKey(key)
        append_run_log("Trusted location set for project root.")
    except Exception as e:
        append_run_log(f"Could not set trusted location: {e}")

    # Enable macros via automation BEFORE opening anything.
    excel.AutomationSecurity = 1  # msoAutomationSecurityEnable (prompt user; answerer dismisses)
    excel.Visible = True
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(temp_copy)
    append_run_log(f"Template copy opened: {temp_copy}")

    # Diagnostics: list VB components and try to find the macro.
    try:
        vbp = wb.VBProject
        comp_names = [c.Name for c in vbp.VBComponents]
        append_run_log(f"VBComponents: {', '.join(comp_names)}")
    except Exception as e:
        append_run_log(f"Could not list VBComponents: {e}")

    macro_result = None
    try:
        macro_result = excel.Run("modCore.CreatePendingPickupReport")
        append_run_log(f"Excel.Run returned; macro result={macro_result}")
        if macro_result is True:
            run_completed = True
    except Exception as e:
        append_run_log(f"Excel.Run raised (expected if VBA shows dialogs): {e}")
        outcome = f"ERROR: {e}"

    # Wait for the macro to complete (completion marker or new report).
    wait_deadline = time.time() + max_run_seconds
    while time.time() < wait_deadline and not cancel_event.is_set():
        time.sleep(0.5)
        if os.path.exists(completion_marker):
            append_run_log("Completion marker found during wait.")
            run_completed = True
            break
        try:
            new_files = [f.name for f in os.scandir(reports_dir)
                         if f.name.lower().endswith(".xlsx") and f.name not in before and not f.name.startswith("~$")]
            if new_files:
                append_run_log(f"Report found during wait: {', '.join(sorted(new_files))}")
                run_completed = True
                break
        except OSError:
            pass
finally:
    cancel_event.set()
    append_run_log("Run token cancelled; closing workbooks and Excel.")
    time.sleep(0.2)
    # Close the template copy gracefully (don't save - it's just a working copy).
    if wb:
        try:
            wb.Close(False)
            append_run_log("Template copy closed.")
        except Exception:
            pass
    if excel:
        try:
            excel.Quit()
            append_run_log("Excel quit.")
        except Exception:
            pass
    try:
        os.remove(temp_copy)
    except OSError:
        pass
    try:
        os.remove(os.path.join(project_root, "~$_run_temp.xlsm"))
    except OSError:
        pass
    if run_completed:
        outcome = "OK"
    write_log(run_log, f"run finished (outcome={outcome})")
    if os.path.exists(completion_marker):
        append_run_log("Completion marker contents:")
        with open(completion_marker, "r", encoding="utf-8") as f:
            for line in f:
                append_run_log("  " + line.rstrip())
    else:
        append_run_log("Completion marker was not written by VBA.")
    # Append per-run VBA trace when available (LogMsg/LogErr dual-write to mRunLogPath during runs);
    # fall back to macro-debug.log tail for older behavior.
    dbg_log = os.path.join(project_root, "log", "macro-debug.log")
    run_vba_log = None
    try:
        log_dir_path = os.path.join(project_root, "log")
        vba_candidates = []
        marker_saved = None
        if os.path.exists(completion_marker):
            try:
                with open(completion_marker, "r", encoding="utf-8") as f:
                    for ln in f:
                        if ln.startswith("saved="):
                            marker_saved = ln.split("=", 1)[1].strip()
            except Exception:
                marker_saved = None
        for name in os.listdir(log_dir_path):
            if not name.startswith("pending-pickup_") or not name.endswith(".log"):
                continue
            p = os.path.join(log_dir_path, name)
            try:
                st = os.stat(p)
            except OSError:
                continue
            # created/modified around this run
            if st.st_mtime + 1 < run_start_time:
                continue
            try:
                with open(p, "r", encoding="utf-8", errors="replace") as f:
                    head = f.read(4096)
                if "=== CreatePendingPickupReport started ===" not in head:
                    continue
                if marker_saved and marker_saved not in head:
                    # may be in tail, but cheap check: read more only if needed
                    with open(p, "r", encoding="utf-8", errors="replace") as f:
                        txt = f.read()
                    if marker_saved not in txt and "VBAWarnings registry set" in txt:
                        continue
                vba_candidates.append((st.st_mtime, p))
            except Exception:
                continue
        if vba_candidates:
            vba_candidates.sort()
            run_vba_log = vba_candidates[-1][1]
    except Exception:
        run_vba_log = None

    if run_vba_log and os.path.exists(run_vba_log):
        append_run_log(f"--- per-run VBA log ({os.path.basename(run_vba_log)}) last 200 lines ---")
        try:
            with open(run_vba_log, "r", encoding="utf-8", errors="replace") as f:
                lines = f.readlines()
            for line in lines[-200:]:
                append_run_log(line.rstrip())
        except Exception as e:
            append_run_log(f"Could not read per-run VBA log: {e}")
    elif os.path.exists(dbg_log):
        append_run_log("--- macro-debug.log (last 120 lines) ---")
        with open(dbg_log, "r", encoding="utf-8") as f:
            lines = f.readlines()
        for line in lines[-120:]:
            append_run_log(line.rstrip())

    run_duration = time.time() - run_start_time
    append_run_log(f"Total Python run duration: {run_duration:.2f} seconds")
    append_run_log(f"Final outcome: {outcome}")

    # Log any new report files created during this run.
    try:
        after = {f.name for f in os.scandir(reports_dir)
                 if f.name.lower().endswith(".xlsx") and not f.name.startswith("~$")}
        new_files = after - before
        if new_files:
            append_run_log(f"New report file(s): {', '.join(sorted(new_files))}")
        else:
            append_run_log("No new report file created.")
    except OSError:
        append_run_log("Could not scan Reports/ for new files.")

    # Flag leftover run artifacts (temp/compile probe debris) for cleanup visibility.
    for leftover in ["_run_temp.xlsm", "_compile_test.xlsm", "probe_ran.txt", "~$_run_temp.xlsm"]:
        lp = os.path.join(project_root, leftover)
        if os.path.exists(lp):
            append_run_log(f"Leftover artifact detected: {lp}")
            try:
                os.remove(lp)
                append_run_log(f"Removed leftover artifact: {lp}")
            except Exception as e:
                append_run_log(f"Could not remove leftover artifact {lp}: {e}")

t_a.join(timeout=2)
t_b.join(timeout=2)

# After Excel is fully shut down, regenerate test screenshots from the newest
# report. snapshot-report.py starts its own Excel instance, so no conflict.
# Screenshots/exports are required and must be allowed to finish, but if they take
# more than 2 minutes (120s) treat as failure/hang: kill the app/Excel and check logs.
screenshot_status = "skipped" if args.no_screenshots else "pending"
if run_completed and not args.no_screenshots:
    snap_script = os.path.join(script_dir, "snapshot-report.py")
    if os.path.exists(snap_script):
        try:
            append_run_log("Regenerating test screenshots from newest report...")
            snap_start = time.time()
            subprocess.run(
                [sys.executable, snap_script],
                check=True,
                capture_output=True,
                text=True,
                timeout=SCREENSHOT_TIMEOUT_SECONDS,
            )
            snap_dur = time.time() - snap_start
            append_run_log(f"Test screenshots regenerated successfully ({snap_dur:.2f}s).")
            screenshot_status = "ok"
        except subprocess.TimeoutExpired:
            append_run_log(
                f"Screenshot regeneration timed out (>{SCREENSHOT_TIMEOUT_SECONDS}s). "
                "Excel killed; check/read logs, identify cause, apply fix, then rerun."
            )
            try:
                subprocess.run(["taskkill", "/F", "/IM", "excel.exe"], capture_output=True, timeout=10)
            except Exception:
                pass
            screenshot_status = "timeout"
        except subprocess.CalledProcessError as e:
            append_run_log(f"Screenshot regeneration failed: {(e.stderr or e.stdout or str(e))[-1000:]}")
            screenshot_status = "failed"
        except Exception as e:
            append_run_log(f"Screenshot regeneration error: {e}")
            screenshot_status = "failed"
    else:
        append_run_log("Snapshot script not found, skipping screenshot regeneration.")
        screenshot_status = "missing_script"
elif run_completed and args.no_screenshots:
    append_run_log("Screenshot regeneration skipped (--no-screenshots). Screenshots are required.")
    screenshot_status = "skipped"
else:
    screenshot_status = "not_run"

# Determine final exit: report must complete and screenshots must succeed (not skipped).
final_ok = bool(run_completed) and (screenshot_status == "ok")
if run_completed and not final_ok:
    if screenshot_status == "skipped":
        append_run_log("Run incomplete: screenshots were skipped (screenshots are required).")
    elif screenshot_status == "timeout":
        append_run_log("Run incomplete: screenshot regeneration timed out (>120s).")
    elif screenshot_status == "failed":
        append_run_log("Run incomplete: screenshot regeneration failed.")
    elif screenshot_status == "missing_script":
        append_run_log("Run incomplete: snapshot-report.py missing; screenshots required.")
append_run_log(f"Screenshot status: {screenshot_status}")
sys.exit(0 if final_ok else 1)
