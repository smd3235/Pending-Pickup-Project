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
SendMessageW = user32.SendMessageW
BM_CLICK = 0x00F5

EnumProc = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)

root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
template = os.path.join(root, "Pending Pickup Template.xlsm")
temp_dir = os.path.join(root, "scripts", "temp")
test_copy = os.path.join(temp_dir, "probe_modules.xlsm")

for p in psutil.process_iter(["name"]):
    if p.info["name"] and p.info["name"].lower() == "excel.exe":
        p.kill()

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
            cls = win_class(hwnd)
            if cls not in ("#32770", "wndclass_desked_gsk"):
                return True
            buttons = []
            texts = [win_text(hwnd)]

            def child(ch, _x):
                ct = win_text(ch)
                if ct:
                    texts.append(ct)
                if win_class(ch) == "Button":
                    buttons.append((ch, ct))
                return True

            EnumChildWindows(hwnd, EnumProc(child), 0)
            if not buttons:
                return True
            text = " || ".join(texts)
            if text in dialogs:
                return True
            dialogs.append(text)
            print("DIALOG:", text, flush=True)
            pick = None
            for pref in ("End", "OK", "Yes", "&OK"):
                for h, t in buttons:
                    if t.replace("&", "").strip().lower() == pref.lower():
                        pick = h
                        break
                if pick:
                    break
            if pick is None:
                pick = buttons[0][0]
            SendMessageW(pick, BM_CLICK, 0, 0)
            return True

        EnumWindows(EnumProc(cb), 0)
        time.sleep(0.3)


modules = ["modDeclarations", "modAudit", "modLogging", "modUtils", "modFormatting", "modDialog", "modCore"]
results = {}

pythoncom.CoInitialize()
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False
wb = None
pid = None
try:
    wb = excel.Workbooks.Open(test_copy)
    pid = wt.DWORD()
    GetWindowThreadProcessId(excel.Hwnd, ctypes.byref(pid))
    pid = pid.value
    print("Excel pid:", pid, flush=True)
    threading.Thread(target=dismisser, args=(pid,), daemon=True).start()

    vbp = wb.VBProject
    for mname in modules:
        comp = vbp.VBComponents(mname)
        cm = comp.CodeModule
        probe = f"zzProbe_{mname}"
        cm.AddFromString(f"\nPublic Sub {probe}()\nEnd Sub\n")
        try:
            excel.Run(probe)
            results[mname] = "OK"
        except Exception as e:
            results[mname] = f"FAILED: {e}"
        print(mname, "->", results[mname], flush=True)
        time.sleep(0.3)
finally:
    stop.set()
    try:
        if wb is not None:
            wb.Close(SaveChanges=False)
    except Exception:
        pass
    try:
        excel.Quit()
    except Exception:
        pass
    pythoncom.CoUninitialize()
    for p in psutil.process_iter(["name"]):
        if p.info["name"] and p.info["name"].lower() == "excel.exe":
            p.kill()

print("\nRESULTS:")
for k, v in results.items():
    print(" ", k, ":", v)
