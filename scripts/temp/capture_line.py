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
IsWindowVisible = user32.IsWindowVisible
SendMessageW = user32.SendMessageW
PostMessageW = user32.PostMessageW
BM_CLICK = 0x00F5
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
VK_RETURN = 13

EnumProc = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)

root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
template = os.path.join(root, "Pending Pickup Template.xlsm")
test_copy = os.path.join(root, "scripts", "temp", "capture_line.xlsm")

for p in psutil.process_iter(["name"]):
    if p.info["name"] and p.info["name"].lower() == "excel.exe":
        p.kill()

if os.path.exists(test_copy):
    os.remove(test_copy)
shutil.copy2(template, test_copy)

stop = threading.Event()
handled = []


def win_text(hwnd):
    n = user32.GetWindowTextLengthW(hwnd)
    b = ctypes.create_unicode_buffer(n + 1)
    user32.GetWindowTextW(hwnd, b, n + 1)
    return b.value


def win_class(hwnd):
    b = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, b, 256)
    return b.value


def click(hwnd):
    SendMessageW(hwnd, BM_CLICK, 0, 0)


def dismisser(pid):
    while not stop.is_set():
        def cb(hwnd, _):
            p = wt.DWORD()
            GetWindowThreadProcessId(hwnd, ctypes.byref(p))
            if p.value != pid or not IsWindowVisible(hwnd):
                return True
            if win_class(hwnd) != "#32770":
                return True
            buttons = []
            texts = [win_text(hwnd)]

            def child(ch, _x):
                t = win_text(ch)
                if t:
                    texts.append(t)
                if win_class(ch) == "Button":
                    buttons.append((ch, t))
                return True

            EnumChildWindows(hwnd, EnumProc(child), 0)
            if not buttons:
                return True
            text = " || ".join(texts)
            key = win_text(hwnd) + "|" + "|".join(t for _, t in buttons)
            if key in handled:
                return True
            handled.append(key)
            print("DIALOG:", text, "buttons:", buttons, flush=True)
            PostMessageW(hwnd, WM_KEYDOWN, VK_RETURN, 0)
            PostMessageW(hwnd, WM_KEYUP, VK_RETURN, 0)
            return True

        EnumWindows(EnumProc(cb), 0)
        time.sleep(0.25)


pythoncom.CoInitialize()
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(test_copy)
pid = wt.DWORD()
GetWindowThreadProcessId(excel.Hwnd, ctypes.byref(pid))
pid = pid.value
print("Excel pid:", pid, flush=True)
threading.Thread(target=dismisser, args=(pid,), daemon=True).start()

err = None
try:
    excel.Run("SetupEmailCC")
    print("run ok", flush=True)
except Exception as e:
    err = e
    print("COM ERROR:", e, flush=True)

time.sleep(1.0)

for attempt in range(10):
    try:
        pane = excel.VBE.ActiveCodePane
        mod = pane.CodeModule.Parent.Name
        sl, sc, el, ec = pane.GetSelection()
        cm = pane.CodeModule
        print("MODULE:", mod, "LINE:", sl, "COL:", sc, flush=True)
        for ln in range(max(1, sl - 3), min(cm.CountOfLines, sl + 3) + 1):
            print(f"  {ln}: {cm.Lines(ln, 1)}", flush=True)
        break
    except Exception as e:
        print("introspect attempt", attempt, "failed:", e, flush=True)
        time.sleep(1.0)

stop.set()
try:
    wb.Close(SaveChanges=False)
except Exception:
    pass
try:
    excel.Quit()
except Exception:
    pass
pythoncom.CoUninitialize()
for p in psutil.process_iter(["name"]):
    if p.info["name"] and p.info["name"].lower() == "excel.exe":
        p.kill()
print("done", flush=True)
