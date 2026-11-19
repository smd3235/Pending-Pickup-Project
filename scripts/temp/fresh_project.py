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
src_dir = os.path.join(root, "vba")
fresh = os.path.join(root, "scripts", "temp", "fresh_project.xlsm")

for p in psutil.process_iter(["name"]):
    if p.info["name"] and p.info["name"].lower() == "excel.exe":
        p.kill()
time.sleep(1)
if os.path.exists(fresh):
    os.remove(fresh)

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
            text = " || ".join(texts)
            if text not in dialogs:
                dialogs.append(text)
            return True

        before = len(dialogs)
        EnumWindows(EnumProc(cb), 0)
        if len(dialogs) > before:
            time.sleep(1.0)
            kill_excel()
            return
        time.sleep(0.25)


pythoncom.CoInitialize()
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False
stop = threading.Event()
try:
    wb = excel.Workbooks.Add()
    wb.SaveAs(fresh, FileFormat=52)
    vbp = wb.VBProject
    for f in sorted(os.listdir(src_dir)):
        if f.endswith(".bas"):
            vbp.VBComponents.Import(os.path.join(src_dir, f))
    print("imported", flush=True)
    pid = wt.DWORD()
    GetWindowThreadProcessId(excel.Hwnd, ctypes.byref(pid))
    threading.Thread(target=watch, args=(pid.value, stop), daemon=True).start()
    for label, name, arg in [("GetLogPath", "GetLogPath", None), ("LogMsg", "LogMsg", "probe")]:
        try:
            if arg is None:
                v = excel.Run(name)
            else:
                v = excel.Run(name, arg)
            print(f"{label}: OK {v!r}", flush=True)
        except Exception as e:
            print(f"{label}: EXC {str(e)[:80]}", flush=True)
        time.sleep(1.0)
    time.sleep(2)
finally:
    stop.set()
    pythoncom.CoUninitialize()
    kill_excel()

print("dialogs:")
for d in dialogs:
    print("  ", d.replace("\n", " ")[:160], flush=True)
