# test-compile.py
# Compile-test harness for Pending_Pickup_Macro.bas
# Copies the template, injects a tiny probe macro, then runs it via COM.
# Running ANY macro forces VBA to compile the ENTIRE project, so a compile
# error anywhere blocks the probe.
#
# Safety:
#  - watch-compile-error.py runs in a separate hidden process; it logs any
#    dialog text to log\watcher.log and auto-dismisses error dialogs.
#  - A 10-second watchdog force-kills this script's own Excel instance.
#    (The whole test normally finishes in a few seconds.)
#  - Excel is always killed/closed at the end.
#
# All output goes to log\compile-test.log with per-line timestamps.

import argparse
import os
import shutil
import subprocess
import sys
import threading
import time
from datetime import datetime

import win32com.client

parser = argparse.ArgumentParser(description="Compile-test Pending_Pickup_Macro.bas")
parser.add_argument("--template", default=None)
args = parser.parse_args()

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)

template = args.template or os.path.join(project_root, "Pending Pickup Template.xlsm")
test_copy = os.path.join(project_root, "_compile_test.xlsm")
marker = os.path.join(project_root, "probe_ran.txt")
watcher = os.path.join(script_dir, "watch-compile-error.py")
log_dir = os.path.join(project_root, "log")
os.makedirs(log_dir, exist_ok=True)
watch_log = os.path.join(log_dir, "watcher.log")
test_log = os.path.join(log_dir, "compile-test.log")


def log(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    with open(test_log, "a", encoding="utf-8") as f:
        f.write(line + "\n")


# 0. Fresh test copy
if os.path.exists(test_copy):
    os.remove(test_copy)
shutil.copy2(template, test_copy)
if os.path.exists(marker):
    os.remove(marker)
log(f"=== Compile test started (template: {template}) ===")

# 1. Start dialog watcher in a separate hidden process
watch_proc = subprocess.Popen(
    [sys.executable, watcher],
    creationflags=subprocess.CREATE_NO_WINDOW,
)
log(f"Dialog watcher started (pid {watch_proc.pid}).")
time.sleep(4)

# 2. Start our own Excel COM instance
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.AutomationSecurity = 1

import psutil
before_pids = {p.pid for p in psutil.process_iter(["name"]) if p.info["name"] and p.info["name"].lower() == "excel.exe"}

# Snapshot pids after starting Excel COM (the COM call above may already have created it)
time.sleep(1)
after_pids = {p.pid for p in psutil.process_iter(["name"]) if p.info["name"] and p.info["name"].lower() == "excel.exe"}
new_pids = after_pids - before_pids
excel_pid = next(iter(new_pids)) if new_pids else None
log(f"Excel COM instance pid(s): {','.join(str(p) for p in new_pids) if new_pids else 'attached to existing instance'}")

# 3. 10-second watchdog: force-kill our Excel pid if anything takes too long
watchdog_timer = None
watchdog_fired = False


def watchdog():
    global watchdog_fired
    time.sleep(10)
    watchdog_fired = True
    try:
        proc = psutil.Process(excel_pid)
        proc.kill()
        log(f"Watchdog killed Excel pid {excel_pid}.")
    except Exception:
        pass


if excel_pid:
    watchdog_timer = threading.Thread(target=watchdog, daemon=True)
    watchdog_timer.start()

# 4. Run the probe
log_mark = 0
if os.path.exists(watch_log):
    with open(watch_log, "r", encoding="utf-8") as f:
        log_mark = sum(1 for _ in f)

run_error = None
wb = None
try:
    wb = excel.Workbooks.Open(test_copy)
    vbp = wb.VBProject
    mod = vbp.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
    mod.Name = "zzCompileProbe"
    probe_code = (
        "Option Explicit\r\n"
        "Public Sub Probe()\r\n"
        f'    Open "{marker}" For Output As #1\r\n'
        '    Print #1, "ran"\r\n'
        "    Close #1\r\n"
        "End Sub"
    )
    mod.CodeModule.AddFromString(probe_code)
    log("Running Probe (forces full-project compile)...")
    excel.Run("Probe")
    log("Run call returned without COM error.")
except Exception as e:
    run_error = str(e)
    log(f"Run call raised: {run_error}")

# 5. Collect watcher log lines produced during this run
time.sleep(2)
new_lines = []
if os.path.exists(watch_log):
    with open(watch_log, "r", encoding="utf-8") as f:
        all_lines = f.readlines()
    if len(all_lines) > log_mark:
        new_lines = all_lines[log_mark:]

compile_ok = os.path.exists(marker)
if compile_ok:
    log("RESULT: COMPILE OK - probe executed, project compiled cleanly.")
else:
    log("RESULT: COMPILE FAILED - probe did not execute.")
    if new_lines:
        log("Captured dialog text (log\\watcher.log):")
        for l in new_lines:
            log("  " + l.rstrip())
    else:
        log(f"No dialog captured. Run error: {run_error}")

# 6. Cleanup
if watchdog_timer and watchdog_timer.is_alive():
    watchdog_timer.join(timeout=1)
try:
    if wb:
        wb.Close(False)
except Exception as e:
    log(f"Close failed: {e}")
try:
    if excel_pid:
        excel.Quit()
        try:
            psutil.Process(excel_pid).kill()
            log(f"Killed lingering Excel pid {excel_pid}.")
        except Exception:
            pass
except Exception as e:
    log(f"Quit failed: {e}")
try:
    watch_proc.kill()
except Exception:
    pass
try:
    os.remove(test_copy)
except OSError:
    pass
log("=== Compile test finished ===")
