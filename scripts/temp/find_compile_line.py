import ctypes
import ctypes.wintypes as wt
import os
import shutil
import threading
import time
import traceback

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
temp_dir = os.path.join(root, "scripts", "temp")
test_copy = os.path.join(temp_dir, "probe_compile_line.xlsm")

for p in psutil.process_iter(["name"]):
    if p.info["name"] and p.info["name"].lower() == "excel.exe":
        p.kill()

if os.path.exists(test_copy):
    os.remove(test_copy)
shutil.copy2(template, test_copy)

dialog_seen = threading.Event()


def win_text(hwnd):
    n = user32.GetWindowTextLengthW(hwnd)
    b = ctypes.create_unicode_buffer(n + 1)
    user32.GetWindowTextW(hwnd, b, n + 1)
    return b.value


def win_class(hwnd):
    b = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, b, 256)
    return b.value


def watcher(pid):
    while not dialog_seen.is_set():
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
            if "Compile error" in text:
                print("DIALOG:", text, flush=True)
                dialog_seen.set()
            return True

        EnumWindows(EnumProc(cb), 0)
        time.sleep(0.3)


pythoncom.CoInitialize()
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(test_copy)
pid = wt.DWORD()
GetWindowThreadProcessId(excel.Hwnd, ctypes.byref(pid))
pid = pid.value
print("Excel pid:", pid, flush=True)

run_error = {}


def runner():
    pythoncom.CoInitialize()
    try:
        excel.Run("SetupEmailCC")
        run_error["result"] = "ok"
    except Exception as e:
        run_error["result"] = str(e)
    pythoncom.CoUninitialize()


t = threading.Thread(target=runner, daemon=True)
t.start()
dialog_seen.wait(timeout=30)

if dialog_seen.is_set():
    for attempt in range(15):
        try:
            pane = excel.VBE.ActiveCodePane
            mod = pane.CodeModule.Parent.Name
            sl, sc, el, ec = pane.GetSelection()
            cm = pane.CodeModule
            print("MODULE:", mod, "LINE:", sl, "COL:", sc, flush=True)
            for ln in range(max(1, sl - 2), min(cm.CountOfLines, sl + 2) + 1):
                print(f"  {ln}: {cm.Lines(ln, 1)}", flush=True)
            break
        except Exception as e:
            print("introspect attempt", attempt, "failed:", e, flush=True)
            time.sleep(1)

# dismiss everything by killing
t.join(timeout=3)
try:
    excel.Quit()
except Exception:
    pass
pythoncom.CoUninitialize()
for p in psutil.process_iter(["name"]):
    if p.info["name"] and p.info["name"].lower() == "excel.exe":
        p.kill()
print("run result:", run_error.get("result"), flush=True)
