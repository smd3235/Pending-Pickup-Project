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
src_dir = os.path.join(root, "vba")
work = os.path.join(root, "scripts", "temp", "decl_probe.xlsm")
temp_bas = os.path.join(root, "scripts", "temp", "modMin3.bas")

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


def probe(label, body):
    dialogs.clear()
    kill_excel()
    time.sleep(1.2)
    for _ in range(10):
        try:
            if os.path.exists(work):
                os.remove(work)
            break
        except (OSError, PermissionError):
            time.sleep(1.0)

    with open(temp_bas, "w", encoding="utf-8", newline="\n") as f:
        f.write('Attribute VB_Name = "modMin"\nOption Explicit\n\n' + body + "\n")

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    stop = threading.Event()
    result = ""
    try:
        wb = excel.Workbooks.Add()
        wb.SaveAs(work, FileFormat=52)
        vbp = wb.VBProject
        vbp.VBComponents.Import(os.path.join(src_dir, "modDeclarations.bas"))
        vbp.VBComponents.Import(temp_bas)
        pid = wt.DWORD()
        GetWindowThreadProcessId(excel.Hwnd, ctypes.byref(pid))
        threading.Thread(target=watch, args=(pid.value, stop), daemon=True).start()
        try:
            v = excel.Run("Probe1")
            result = f"OK {v!r}"
        except Exception as e:
            result = f"EXC {str(e)[:90]}"
        time.sleep(1.2)
    finally:
        stop.set()
        pythoncom.CoUninitialize()
        kill_excel()
    print(f"{label:26s} {result}", flush=True)
    for d in dialogs:
        print("    ", d.replace("\n", " ")[:130], flush=True)


probe("plain", "Public Sub Probe1()\n    Debug.Print \"hi\"\nEnd Sub")
probe("GetLogPath", "Public Function Probe1() As String\n    Probe1 = GetLogPath()\nEnd Function")
probe("mInReportRun", "Public Function Probe1() As Boolean\n    Probe1 = mInReportRun\nEnd Function")
probe("mRunLogPath", "Public Function Probe1() As String\n    Probe1 = mRunLogPath\nEnd Function")
probe("mAuditIssues", "Public Sub Probe1()\n    Set mAuditIssues = New Collection\nEnd Sub")
probe("mAuditStartTime", "Public Function Probe1() As Date\n    Probe1 = mAuditStartTime\nEnd Function")
