import ctypes
import ctypes.wintypes as wt
import os
import re
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
work = os.path.join(root, "scripts", "temp", "log_bisect.xlsm")
src = open(os.path.join(src_dir, "modLogging.bas"), encoding="utf-8").read().replace("\r\n", "\n")
lines = src.split("\n")
start_re = re.compile(r"^(?:Public\s+|Private\s+)?(?:Sub|Function)\s+\w+")
starts = [i for i, ln in enumerate(lines) if start_re.match(ln)]
header = "\n".join(lines[: starts[0]]).rstrip()
blocks = {}
order = []
for idx, s in enumerate(starts):
    e = starts[idx + 1] if idx + 1 < len(starts) else len(lines)
    chunk = "\n".join(lines[s:e]).rstrip()
    nm = re.search(r"(?:Sub|Function)\s+(\w+)", chunk).group(1)
    blocks[nm] = chunk
    order.append(nm)

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


def test(names, entry, arg):
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

    logging_code = header + "\n\n" + "\n\n".join(blocks[n] for n in names) + "\n"
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
        mod = vbp.VBComponents.Add(1)
        mod.Name = "modLoggingTest"
        mod.CodeModule.AddFromString(logging_code)
        pid = wt.DWORD()
        GetWindowThreadProcessId(excel.Hwnd, ctypes.byref(pid))
        threading.Thread(target=watch, args=(pid.value, stop), daemon=True).start()
        try:
            if arg is None:
                excel.Run("modLoggingTest." + entry)
            else:
                excel.Run("modLoggingTest." + entry, arg)
            print(f"  {names} -> ran", flush=True)
        except Exception as e:
            print(f"  {names} -> EXC {str(e)[:400]}", flush=True)
        time.sleep(1.5)
    finally:
        stop.set()
        pythoncom.CoUninitialize()
        kill_excel()
    bad = any("Compile error" in d for d in dialogs)
    print(f"  {names}: compile_error={bad}", flush=True)
    return bad


for combo in (["LogMsg"], ["LogErr"], ["WriteCompletionMarker"], ["BuildRunLogPath"], order):
    test(combo, "LogMsg", "probe")
