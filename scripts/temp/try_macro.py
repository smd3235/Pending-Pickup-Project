import ctypes
import ctypes.wintypes as wt
import os
import shutil
import sys
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
test_copy = os.path.join(root, "scripts", "temp", "try_macro.xlsm")

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
            return True

        before = len(dialogs)
        EnumWindows(EnumProc(cb), 0)
        if len(dialogs) > before:
            time.sleep(1.2)
            kill_excel()
            return
        time.sleep(0.3)


def kill_excel():
    for p in psutil.process_iter(["name"]):
        if p.info["name"] and p.info["name"].lower() == "excel.exe":
            p.kill()


def run_one(macro):
    dialogs.clear()
    kill_excel()
    time.sleep(2.0)
    for i in range(10):
        try:
            if os.path.exists(test_copy):
                os.remove(test_copy)
            shutil.copy2(template, test_copy)
            break
        except (OSError, PermissionError):
            time.sleep(1.0)

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    stop = threading.Event()
    result = ""
    try:
        wb = excel.Workbooks.Open(test_copy)
        pid = wt.DWORD()
        GetWindowThreadProcessId(excel.Hwnd, ctypes.byref(pid))
        threading.Thread(target=watch, args=(pid.value, stop), daemon=True).start()
        try:
            if "|" in macro:
                name, arg = macro.split("|", 1)
                value = excel.Run(name, arg)
            else:
                value = excel.Run(macro)
            result = f"OK ({value!r})"
        except Exception as e:
            result = f"EXC {str(e)[:80]}"
        time.sleep(6)
    finally:
        stop.set()
        pythoncom.CoUninitialize()
        kill_excel()
    print(f"{macro:32s} -> {result}", flush=True)
    for d in dialogs:
        print("    DIALOG:", d.replace("\n", " ")[:150], flush=True)


for m in sys.argv[1:]:
    run_one(m)
