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
src_dir = os.path.join(root, "vba")
work = os.path.join(root, "scripts", "temp", "log_bisect.xlsm")
temp_bas = os.path.join(root, "scripts", "temp", "modLogging_variant.bas")

src = open(os.path.join(src_dir, "modLogging.bas"), encoding="utf-8").read().replace("\r\n", "\n")
lines = src.split("\n")
start_re = re.compile(r"^(?:Public\s+|Private\s+)?(?:Sub|Function)\s+\w+")
starts = [i for i, ln in enumerate(lines) if start_re.match(ln)]
s0 = starts[0]
s1 = starts[1]
header = "\n".join(lines[:s0]).rstrip()
logmsg_lines = lines[s0:s1]
rest = "\n".join(lines[s1:]).rstrip()

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


def test(label, body_lines):
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

    code = header + "\n\n" + "\n".join(body_lines) + "\n\n" + rest + "\n"
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
        except Exception as e:
            pass
        time.sleep(1.5)
    finally:
        stop.set()
        pythoncom.CoUninitialize()
        kill_excel()
    bad = any("Compile error" in d for d in dialogs)
    print(f"{label:34s} compile_error={bad}", flush=True)
    return bad


print("LogMsg block:")
for i, ln in enumerate(logmsg_lines):
    print(f"{i:2d}: {ln}")

print()
test("baseline", logmsg_lines)

# Variant: drop the mInReportRun block (from 'If mInReportRun' to matching End If before End Sub)
idx = next(i for i, ln in enumerate(logmsg_lines) if "mInReportRun" in ln)
tail = logmsg_lines[-1]
test("no mInReportRun block", logmsg_lines[:idx] + [tail])

# Variant: drop the MkDir/Dir line
idx2 = next(i for i, ln in enumerate(logmsg_lines) if "MkDir" in ln)
test("no Dir/MkDir line", [ln for i, ln in enumerate(logmsg_lines) if i != idx2])

# Variant: drop Open GetLogPath block
idx3 = next(i for i, ln in enumerate(logmsg_lines) if "Open GetLogPath" in ln)
body = logmsg_lines[:idx3] + [logmsg_lines[-1]]
test("only declarations/Format", body)
