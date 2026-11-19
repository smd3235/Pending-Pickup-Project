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
temp = os.path.join(root, "scripts", "temp")

src = open(os.path.join(root, "vba", "modDeclarations.bas"), encoding="utf-8").read().replace("\r\n", "\n")
body = src.split("Option Explicit", 1)[1].lstrip("\n")

lines = body.split("\n")
const_lines = [ln for ln in lines if ln.startswith("Public Const")]
var_lines = [ln for ln in lines if re.match(r"Public m\w+ As ", ln)]

func_names = re.findall(r"^(Public Function (\w+))", body, re.M)
func_blocks = []
idx_starts = [i for i, ln in enumerate(lines) if ln.startswith("Public Function")]
for j, s in enumerate(idx_starts):
    e = idx_starts[j + 1] if j + 1 < len(idx_starts) else len(lines)
    func_blocks.append("\n".join(lines[s:e]).rstrip())

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


def test(idx, label, decl_body):
    dialogs.clear()
    kill_excel()
    time.sleep(1.5)
    work = os.path.join(temp, f"declsec_{idx}.xlsm")
    decl = os.path.join(temp, f"modDeclSec_{idx}.bas")
    probe = os.path.join(temp, f"modProbeSec_{idx}.bas")
    with open(decl, "w", encoding="utf-8", newline="\n") as f:
        f.write('Attribute VB_Name = "modDeclSec"\nOption Explicit\n\n' + decl_body + "\n\nPublic mInReportRun As Boolean\n")
    with open(probe, "w", encoding="utf-8", newline="\n") as f:
        f.write('Attribute VB_Name = "modProbeSec"\nOption Explicit\n\nPublic Function Probe1() As Boolean\n    Probe1 = mInReportRun\nEnd Function\n')

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
        vbp.VBComponents.Import(decl)
        vbp.VBComponents.Import(probe)
        pid = wt.DWORD()
        GetWindowThreadProcessId(excel.Hwnd, ctypes.byref(pid))
        threading.Thread(target=watch, args=(pid.value, stop), daemon=True).start()
        try:
            v = excel.Run("Probe1")
            out = f"OK {v!r}"
        except Exception as e:
            out = "EXC " + str(e)[:60]
        time.sleep(1.2)
    finally:
        stop.set()
        pythoncom.CoUninitialize()
        kill_excel()
    bad = any("Compile error" in d for d in dialogs)
    print(f"{label:26s} {out}  compile_error={bad}", flush=True)
    return bad


cases = [(f"func {i}", blk) for i, blk in enumerate(func_blocks)]
cases.append(("all funcs", "\n\n".join(func_blocks)))
for i, (label, code) in enumerate(cases):
    test(i, label, code)
