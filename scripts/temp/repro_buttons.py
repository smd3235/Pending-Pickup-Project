import ctypes
import ctypes.wintypes as wt
import os
import shutil
import sys
import threading
import time

import pythoncom
import win32com.client

user32 = ctypes.windll.user32
EnumWindows = user32.EnumWindows
EnumChildWindows = user32.EnumChildWindows
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
GetWindowTextW = user32.GetWindowTextW
GetClassNameW = user32.GetClassNameW
PostMessageW = user32.PostMessageW
IsWindowVisible = user32.IsWindowVisible
WM_COMMAND = 0x0111
IDOK = 1

EnumProc = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)

root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
template = os.path.join(root, "Pending Pickup Template.xlsm")
temp_dir = os.path.join(root, "scripts", "temp")
test_copy = os.path.join(temp_dir, "repro_buttons.xlsm")
log_dir = os.path.join(os.environ["TEMP"], "PendingPickupReport", "log")
os.makedirs(log_dir, exist_ok=True)

with open(os.path.join(log_dir, "auto-source.txt"), "w", encoding="utf-8") as f:
    f.write(os.path.join(root, "AllOpenTreatmentPlansResults618.xlsx"))
with open(os.path.join(log_dir, "auto-skip-prior.txt"), "w", encoding="utf-8") as f:
    f.write("1")

if os.path.exists(test_copy):
    os.remove(test_copy)
shutil.copy2(template, test_copy)

dialogs = []
stop = threading.Event()


def win_text(hwnd):
    n = user32.GetWindowTextLengthW(hwnd)
    b = ctypes.create_unicode_buffer(n + 1)
    user32.GetWindowTextW(hwnd, b, n + 1)
    return b.value


def win_class(hwnd):
    b = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, b, 256)
    return b.value


def dismisser(pid):
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
            text = " || ".join(texts)
            if text not in dialogs:
                dialogs.append(text)
                print("DIALOG:", text, flush=True)
            PostMessageW(hwnd, WM_COMMAND, IDOK, 0)
            return True

        EnumWindows(EnumProc(cb), 0)
        time.sleep(0.4)


pythoncom.CoInitialize()
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False
try:
    wb = excel.Workbooks.Open(test_copy)
    pid = None
    import psutil
    hwnd = excel.Hwnd
    pid = wt.DWORD()
    GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    pid = pid.value
    print("Excel pid:", pid, flush=True)
    th = threading.Thread(target=dismisser, args=(pid,), daemon=True)
    th.start()

    for macro in ["SetupEmailCC", "SetupDirectorEmail", "CreatePendingPickupReport"]:
        print(f"--- running {macro}", flush=True)
        try:
            result = excel.Run(macro)
            print(f"  returned: {result!r}", flush=True)
        except Exception as e:
            print(f"  COM ERROR: {e}", flush=True)
        time.sleep(0.5)

    wb.Close(SaveChanges=False)
finally:
    stop.set()
    try:
        excel.Quit()
    except Exception:
        pass
    pythoncom.CoUninitialize()

print("dialogs seen:", len(dialogs))
