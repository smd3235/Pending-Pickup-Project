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


def test(label, body):
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

    code = (
        'Attribute VB_Name = "modMin"\n'
        "Option Explicit\n\n"
        "Public Sub LogMsg(ByVal msg As String)\n"
        + body
        + "\nEnd Sub\n"
    )
    with open(temp_bas, "w", encoding="utf-8", newline="\n") as f:
        f.write(code)

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    stop = threading.Event()
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
            excel.Run("LogMsg", "probe")
        except Exception:
            pass
        time.sleep(1.5)
    finally:
        stop.set()
        pythoncom.CoUninitialize()
        kill_excel()
    bad = any("Compile error" in d for d in dialogs)
    print(f"{label:24s} compile_error={bad}", flush=True)
    for d in dialogs:
        print("    ", d.replace("\n", " ")[:140], flush=True)
    return bad


cases = {
    "freefile": "    Dim f As Integer\n    f = FreeFile",
    "errclear": "    On Error Resume Next\n    Err.Clear\n    If Err.Number <> 0 Then Exit Sub",
    "dirmkdir": r'''    If Len(Dir$(Left$(GetLogPath(), InStrRev(GetLogPath(), "\")), vbDirectory)) = 0 Then MkDir Left$(GetLogPath(), InStrRev(GetLogPath(), "\"))''',
    "openprint": '    Dim f As Integer\n    Dim logLine As String\n    f = FreeFile\n    Open GetLogPath() For Append As #f\n    Print #f, logLine\n    Close #f',
    "runlogblock": "    If mInReportRun And Len(mRunLogPath) > 0 Then\n        Debug.Print mRunLogPath\n    End If",
}

for label, body in cases.items():
    test(label, body)
