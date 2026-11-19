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
work = os.path.join(root, "scripts", "temp", "min_bisect.xlsm")
temp_bas = os.path.join(root, "scripts", "temp", "modMin.bas")

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


def test(label, code, imports_decl=True):
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
        f.write('Attribute VB_Name = "modMin"\n' + code)

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    stop = threading.Event()
    try:
        wb = excel.Workbooks.Add()
        wb.SaveAs(work, FileFormat=52)
        vbp = wb.VBProject
        if imports_decl:
            vbp.VBComponents.Import(os.path.join(src_dir, "modDeclarations.bas"))
        vbp.VBComponents.Import(temp_bas)
        pid = wt.DWORD()
        GetWindowThreadProcessId(excel.Hwnd, ctypes.byref(pid))
        threading.Thread(target=watch, args=(pid.value, stop), daemon=True).start()
        try:
            excel.Run("LogMsg", "probe")
        except Exception as e:
            pass
        time.sleep(1.5)
    finally:
        stop.set()
        pythoncom.CoUninitialize()
        kill_excel()
    bad = any("Compile error" in d for d in dialogs)
    print(f"{label:40s} compile_error={bad}", flush=True)
    return bad


test("empty sub", 'Option Explicit\n\nPublic Sub LogMsg(ByVal msg As String)\nEnd Sub\n')
test("with dims", 'Option Explicit\n\nPublic Sub LogMsg(ByVal msg As String)\n    Dim f As Integer\n    Dim logLine As String\nEnd Sub\n')
test("with format", 'Option Explicit\n\nPublic Sub LogMsg(ByVal msg As String)\n    Dim logLine As String\n    logLine = Format(Now, "yyyy-mm-dd hh:n:ss") & " | " & msg\nEnd Sub\n')
test("no modDeclarations import", 'Option Explicit\n\nPublic Sub LogMsg(ByVal msg As String)\nEnd Sub\n', imports_decl=False)
