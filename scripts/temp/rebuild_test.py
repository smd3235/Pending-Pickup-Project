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
src_dir = os.path.join(root, "vba")
work = os.path.join(root, "scripts", "temp", "rebuilt.xlsm")

for p in psutil.process_iter(["name"]):
    if p.info["name"] and p.info["name"].lower() == "excel.exe":
        p.kill()
if os.path.exists(work):
    os.remove(work)
shutil.copy2(template, work)

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


def watcher(pid):
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
            return True

        EnumWindows(EnumProc(cb), 0)
        time.sleep(0.4)


pythoncom.CoInitialize()
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False
try:
    wb = excel.Workbooks.Open(work)
    vbp = wb.VBProject
    for comp in list(vbp.VBComponents):
        if comp.Type == 1:
            vbp.VBComponents.Remove(comp)
    for f in sorted(os.listdir(src_dir)):
        if f.endswith(".bas"):
            vbp.VBComponents.Import(os.path.join(src_dir, f))
    print("imported modules:", [c.Name for c in vbp.VBComponents if c.Type == 1], flush=True)
    wb.Save()
    wb.Close(SaveChanges=False)
    excel.Quit()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(work)
    pid = wt.DWORD()
    GetWindowThreadProcessId(excel.Hwnd, ctypes.byref(pid))
    pid = pid.value
    threading.Thread(target=watcher, args=(pid,), daemon=True).start()
    try:
        excel.Run("SetupEmailCC")
        print("RESULT: SetupEmailCC ran with no COM error", flush=True)
    except Exception as e:
        print("RESULT: COM ERROR", str(e)[:150], flush=True)
    time.sleep(2)
finally:
    stop.set()
    try:
        excel.Quit()
    except Exception:
        pass
    pythoncom.CoUninitialize()
    for p in psutil.process_iter(["name"]):
        if p.info["name"] and p.info["name"].lower() == "excel.exe":
            p.kill()

print("dialogs:", len(dialogs), flush=True)
