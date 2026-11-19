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
template = os.path.join(root, "Pending Pickup Template.xlsm")
temp_dir = os.path.join(root, "scripts", "temp")
test_copy = os.path.join(temp_dir, "bisect.xlsm")
src = open(os.path.join(root, "vba", "modDialog.bas"), encoding="utf-8").read().replace("\r\n", "\n")
src_lines = src.split("\n")

start_re = re.compile(r"^(?:Public\s+|Private\s+|Friend\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+\w+")
starts = [i for i, ln in enumerate(src_lines) if start_re.match(ln)]
header = "\n".join(src_lines[: starts[0]]).rstrip()
blocks = {}
order = []
for idx, s in enumerate(starts):
    e = starts[idx + 1] if idx + 1 < len(starts) else len(src_lines)
    chunk = "\n".join(src_lines[s:e]).rstrip()
    m = re.match(r"^(?:Public\s+|Private\s+|Friend\s+)?(?:Sub|Function|Property\s+\w+\s+)?(\w+)", chunk, re.I)
    name = m.group(1) if m else f"block{idx}"
    # match signature line explicitly
    sig = re.search(r"(?:Sub|Function)\s+(\w+)", chunk)
    name = sig.group(1) if sig else name
    blocks[name] = chunk
    order.append(name)

print("procedures:", order, flush=True)

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
    hit = False
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


def test_subset(names):
    dialogs.clear()
    kill_excel()
    time.sleep(1.2)
    for _ in range(10):
        try:
            if os.path.exists(test_copy):
                os.remove(test_copy)
            shutil.copy2(template, test_copy)
            break
        except (OSError, PermissionError):
            time.sleep(1.0)
    code = header + "\n\n" + "\n\n".join(blocks[n] for n in names) + "\n"
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    stop = threading.Event()
    try:
        wb = excel.Workbooks.Open(test_copy)
        cm = wb.VBProject.VBComponents("modDialog").CodeModule
        cm.DeleteLines(1, cm.CountOfLines)
        cm.AddFromString(code)
        pid = wt.DWORD()
        GetWindowThreadProcessId(excel.Hwnd, ctypes.byref(pid))
        threading.Thread(target=watch, args=(pid.value, stop), daemon=True).start()
        try:
            excel.Run("modDialog.GetEmailTo")
        except Exception:
            pass
        time.sleep(1.5)
    finally:
        stop.set()
        pythoncom.CoUninitialize()
        kill_excel()
    compile_error = any("Compile error" in d for d in dialogs)
    return compile_error


if not test_subset(order):
    print("full set compiles clean - cannot bisect", flush=True)
    raise SystemExit

candidates = list(order)
probe = "GetEmailTo"
if probe in candidates:
    candidates.remove(probe)
while len(candidates) > 1:
    mid = len(candidates) // 2
    first = candidates[:mid]
    subset = first + [probe]
    bad = test_subset(subset)
    print(f"test {len(subset)} procs ({first[0]}..{first[-1]}): compile_error={bad}", flush=True)
    if bad:
        candidates = first
    else:
        candidates = candidates[mid:]

print("SUSPECT:", candidates, flush=True)
if candidates:
    print(blocks[candidates[0]], flush=True)
