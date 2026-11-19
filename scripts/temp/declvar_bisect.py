import ctypes
import ctypes.wintypes as wt
import os
import re
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
work = os.path.join(root, "scripts", "temp", "declvar.xlsm")
mod_decl = os.path.join(root, "scripts", "temp", "modDeclTest.bas")
mod_probe = os.path.join(root, "scripts", "temp", "modProbeTest.bas")

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


with open(mod_probe, "w", encoding="utf-8", newline="\n") as f:
    f.write('Attribute VB_Name = "modProbe"\nOption Explicit\n\nPublic Function Probe1() As Boolean\n    Probe1 = mInReportRun\nEnd Function\n')

src = open(os.path.join(root, "vba", "modDeclarations.bas"), encoding="utf-8").read().replace("\r\n", "\n")
body = src.split("Option Explicit", 1)[1].lstrip("\n")
var_block = "\n".join(ln for ln in body.split("\n") if ln.startswith("Public m"))
const_block = "\n".join(ln for ln in body.split("\n") if ln.startswith("Public Const"))
logdir_fn = "Public Function GetLogPath() As String\n    GetLogPath = \"C:\\Temp\\x.log\"\nEnd Function"

variants = {
    "public var only": "Public mInReportRun As Boolean",
    "var block": var_block,
    "const+var": const_block + "\n\n" + var_block,
    "var+logpath fn": var_block + "\n\n" + logdir_fn,
    "plain var + fn": "Public mInReportRun As Boolean\n\n" + logdir_fn,
    "full body": body,
}

for label, text in variants.items():
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
    with open(mod_decl, "w", encoding="utf-8", newline="\n") as f:
        f.write('Attribute VB_Name = "modDeclTest"\nOption Explicit\n\n' + text + "\n")
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    stop = threading.Event()
    out = ""
    try:
        wb = excel.Workbooks.Add()
        wb.SaveAs(work, FileFormat=52)
        vbp = wb.VBProject
        vbp.VBComponents.Import(mod_decl)
        vbp.VBComponents.Import(mod_probe)
        pid = wt.DWORD()
        GetWindowThreadProcessId(excel.Hwnd, ctypes.byref(pid))
        threading.Thread(target=watch, args=(pid.value, stop), daemon=True).start()
        try:
            v = excel.Run("Probe1")
            out = f"OK {v!r}"
        except Exception as e:
            out = "EXC " + str(e)[:80]
        time.sleep(1.2)
    finally:
        stop.set()
        pythoncom.CoUninitialize()
        kill_excel()
    bad = any("Compile error" in d for d in dialogs)
    print(f"{label:18s} {out}  compile_error={bad}", flush=True)
