import ctypes
import ctypes.wintypes as wt
import os
import shutil
import threading
import time

import psutil
import pythoncom
import win32com.client

user32 = ctypes.windll.user32
EnumWindows = user32.EnumWindows
EnumChildWindows = user32.EnumChildWindows
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
GetWindowTextW = user32.GetWindowTextW
GetClassNameW = user32.GetClassNameW
IsWindowVisible = user32.IsWindowVisible

EnumProc = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)

root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
template = os.path.join(root, "Pending Pickup Template.xlsm")
reports_dir = os.path.join(root, "Reports")
work = os.path.join(root, "scripts", "temp", "run_check.xlsm")
log_dir = os.path.join(os.environ["TEMP"], "PendingPickupReport", "log")
os.makedirs(log_dir, exist_ok=True)
with open(os.path.join(log_dir, "auto-source.txt"), "w", encoding="utf-8") as f:
    f.write(os.path.join(root, "AllOpenTreatmentPlansResults618.xlsx"))
with open(os.path.join(log_dir, "auto-skip-prior.txt"), "w", encoding="utf-8") as f:
    f.write("1")

dialogs = []


def win_text(hwnd):
    n = user32.GetWindowTextLengthW(hwnd)
    b = ctypes.create_unicode_buffer(n + 1)
    user32.GetWindowTextW(hwnd, b, n + 1)
    return b.value


def win_class(hwnd):
    b = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, b, 256)
    return b.value


def kill_excel():
    for p in psutil.process_iter(["name"]):
        if p.info["name"] and p.info["name"].lower() == "excel.exe":
            p.kill()


def watch(pid, stop):
    while not stop.is_set():
        def cb(hwnd, _):
            p = wt.DWORD()
            GetWindowThreadProcessId(hwnd, ctypes.byref(p))
            if p.value != pid or not IsWindowVisible(hwnd):
                return True
            if win_class(hwnd) != "#32770":
                return True
            texts = [win_text(hwnd)]

            def child(ch, _x):
                t = win_text(ch)
                if t:
                    texts.append(t)
                return True

            EnumChildWindows(hwnd, EnumProc(child), 0)
            t = " || ".join(texts)
            if t not in dialogs:
                dialogs.append(t)
                print("DIALOG:", t.replace("\n", " ")[:180], flush=True)
            return True

        EnumWindows(EnumProc(cb), 0)
        time.sleep(0.4)


kill_excel()
time.sleep(1.0)
if os.path.exists(work):
    os.remove(work)
shutil.copy2(template, work)
before = set(os.listdir(reports_dir))

pythoncom.CoInitialize()
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False
stop = threading.Event()
result = ""
try:
    wb = excel.Workbooks.Open(work)
    pid = wt.DWORD()
    GetWindowThreadProcessId(excel.Hwnd, ctypes.byref(pid))
    threading.Thread(target=watch, args=(pid.value, stop), daemon=True).start()
    try:
        v = excel.Run("CreatePendingPickupReport")
        result = f"OK {v!r}"
    except Exception as e:
        result = "EXC " + str(e)[:120]
    deadline = time.time() + 40
    while time.time() < deadline:
        new = set(os.listdir(reports_dir)) - before
        if new:
            break
        time.sleep(1)
finally:
    stop.set()
    time.sleep(0.5)
    try:
        wb.Close(SaveChanges=False)
    except Exception:
        pass
    try:
        excel.Quit()
    except Exception:
        pass
    pythoncom.CoUninitialize()
    kill_excel()

print("RESULT:", result)
new_files = sorted(set(os.listdir(reports_dir)) - before)
print("NEW REPORTS:", new_files)
