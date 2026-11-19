import ctypes
import ctypes.wintypes as wt
import os
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
work = os.path.join(root, "scripts", "temp", "pubvar.xlsm")
bas_a = os.path.join(root, "scripts", "temp", "modPubA.bas")
bas_b = os.path.join(root, "scripts", "temp", "modPubB.bas")

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
            return True

        before = len(dialogs)
        EnumWindows(EnumProc(cb), 0)
        if len(dialogs) > before:
            time.sleep(1.0)
            kill_excel()
            return
        time.sleep(0.25)


with open(bas_a, "w", encoding="utf-8", newline="\n") as f:
    f.write('Attribute VB_Name = "modPubA"\nOption Explicit\n\nPublic gTestFlag As Boolean\nPublic gTestName As String\nPublic gTestList As Collection\n')
with open(bas_b, "w", encoding="utf-8", newline="\n") as f:
    f.write('Attribute VB_Name = "modPubB"\nOption Explicit\n\nPublic Function Probe1() As Boolean\n    gTestFlag = True\n    gTestName = "x"\n    Set gTestList = New Collection\n    Probe1 = gTestFlag\nEnd Function\n')

kill_excel()
time.sleep(1)
if os.path.exists(work):
    os.remove(work)

pythoncom.CoInitialize()
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False
stop = threading.Event()
try:
    wb = excel.Workbooks.Add()
    wb.SaveAs(work, FileFormat=52)
    vbp = wb.VBProject
    vbp.VBComponents.Import(bas_a)
    vbp.VBComponents.Import(bas_b)
    pid = wt.DWORD()
    GetWindowThreadProcessId(excel.Hwnd, ctypes.byref(pid))
    threading.Thread(target=watch, args=(pid.value, stop), daemon=True).start()
    try:
        v = excel.Run("Probe1")
        print("OK", v, flush=True)
    except Exception as e:
        print("EXC", str(e)[:150], flush=True)
    time.sleep(1.5)
finally:
    stop.set()
    pythoncom.CoUninitialize()
    kill_excel()
for d in dialogs:
    print("DIALOG:", d.replace("\n", " ")[:150], flush=True)
