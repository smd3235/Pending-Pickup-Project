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
PostMessageW = user32.PostMessageW
WM_COMMAND = 0x0111
IDOK = 1

EnumProc = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)

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
    seen = []
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
            if text not in seen:
                seen.append(text)
                print("DIALOG:", text, flush=True)
            PostMessageW(hwnd, WM_COMMAND, IDOK, 0)
            return True

        EnumWindows(EnumProc(cb), 0)
        time.sleep(0.3)


template = r"C:\Users\smd32\Downloads\Pending Pickup Project\Pending Pickup Template.xlsm"
temp_dir = r"C:\Users\smd32\Downloads\Pending Pickup Project\scripts\temp"

for p in psutil.process_iter(["name"]):
    if p.info["name"] and p.info["name"].lower() == "excel.exe":
        p.kill()

pythoncom.CoInitialize()
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False
try:
    wb = excel.Workbooks.Add()
    pid = wt.DWORD()
    GetWindowThreadProcessId(excel.Hwnd, ctypes.byref(pid))
    pid = pid.value
    threading.Thread(target=dismisser, args=(pid,), daemon=True).start()

    blank = os.path.join(temp_dir, "blank_probe.xlsm")
    wb.SaveAs(blank, FileFormat=52)
    mod = wb.VBProject.VBComponents.Add(1)
    mod.Name = "zzBlank"
    mod.CodeModule.AddFromString("Public Sub zzHello()\nEnd Sub\n")
    try:
        excel.Run("zzHello")
        print("blank workbook macro: OK", flush=True)
    except Exception as e:
        print("blank workbook macro: FAILED", e, flush=True)

    print("--- now open template and just call a trivial existing function", flush=True)
    wb.Close(SaveChanges=False)
    twb = excel.Workbooks.Open(template)
    for macro in ["GetLogPath", "SetupEmailCC"]:
        try:
            r = excel.Run(macro)
            print(macro, "OK", r, flush=True)
        except Exception as e:
            print(macro, "FAILED", e, flush=True)
    try:
        twb.Close(SaveChanges=False)
    except Exception:
        pass
finally:
    stop.set()
    time.sleep(0.5)
    try:
        excel.Quit()
    except Exception:
        pass
    pythoncom.CoUninitialize()
    for p in psutil.process_iter(["name"]):
        if p.info["name"] and p.info["name"].lower() == "excel.exe":
            p.kill()
