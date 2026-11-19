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
PostMessageW = user32.PostMessageW
IsWindowVisible = user32.IsWindowVisible
WM_COMMAND = 0x0111
IDOK = 1
IDCANCEL = 2

EnumProc = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)

root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
template = os.path.join(root, "Pending Pickup Template.xlsm")
temp_dir = os.path.join(root, "scripts", "temp")
test_copy = os.path.join(temp_dir, "repro_compile.xlsm")
for p in psutil.process_iter(["name"]):
    if p.info["name"] and p.info["name"].lower() == "excel.exe":
        p.kill()

if os.path.exists(test_copy):
    os.remove(test_copy)
shutil.copy2(template, test_copy)

dialogs = []
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


def button_ids(hwnd):
    ids = []

    def child(ch, _x):
        cls = win_class(ch)
        if cls == "Button":
            ids.append((ch, win_text(ch)))
        return True

    EnumChildWindows(hwnd, EnumProc(child), 0)
    return ids


def dismisser(pid):
    seen_compile = False
    while not stop.is_set():
        def cb(hwnd, _):
            nonlocal seen_compile
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
            if text in dialogs:
                return True
            dialogs.append(text)
            print("DIALOG:", text, flush=True)
            if "Compile error" in text:
                seen_compile = True
                PostMessageW(hwnd, WM_COMMAND, IDOK, 0)
            elif seen_compile:
                ids = button_ids(hwnd)
                print("  buttons:", ids, flush=True)
                PostMessageW(hwnd, WM_COMMAND, IDCANCEL, 0)
            else:
                PostMessageW(hwnd, WM_COMMAND, IDOK, 0)
            return True

        EnumWindows(EnumProc(cb), 0)
        time.sleep(0.4)


pythoncom.CoInitialize()
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False
pid = None
try:
    wb = excel.Workbooks.Open(test_copy)
    pid = wt.DWORD()
    GetWindowThreadProcessId(excel.Hwnd, ctypes.byref(pid))
    pid = pid.value
    print("Excel pid:", pid, flush=True)
    threading.Thread(target=dismisser, args=(pid,), daemon=True).start()

    try:
        excel.Run("SetupEmailCC")
        print("run ok", flush=True)
    except Exception as e:
        print("COM ERROR:", e, flush=True)

    time.sleep(1.5)

    try:
        vbe = excel.VBE
        pane = vbe.ActiveCodePane
        print("active code pane module:", pane.CodeModule.Parent.Name, flush=True)
        start_line, start_col, end_line, end_col = pane.GetSelection()
        print("selection:", start_line, start_col, end_line, end_col, flush=True)
        mod = pane.CodeModule
        for ln in range(max(1, start_line - 3), min(mod.CountOfLines, start_line + 3) + 1):
            print(f"  {ln}: {mod.Lines(ln, 1)}", flush=True)
    except Exception as e:
        print("VBE introspection failed:", e, flush=True)

    print("total dialogs:", len(dialogs), flush=True)
    try:
        wb.Close(SaveChanges=False)
    except Exception as e:
        print("close failed:", e, flush=True)
finally:
    stop.set()
    try:
        excel.Quit()
    except Exception:
        pass
    pythoncom.CoUninitialize()
    for p in psutil.process_iter(["name"]):
        if p.info["name"] and p.info["name"].lower() == "excel.exe":
            p.kill()
