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
GetWindowRect = user32.GetWindowRect
SetForegroundWindow = user32.SetForegroundWindow
SetCursorPos = user32.SetCursorPos
GetCursorPos = user32.GetCursorPos
mouse_event = user32.mouse_event
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004

EnumProc = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)

root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
template = os.path.join(root, "Pending Pickup Template.xlsm")
test_copy = os.path.join(root, "scripts", "temp", "line_hunt2.xlsm")

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


def click_control(hwnd):
    rect = wt.RECT()
    GetWindowRect(hwnd, ctypes.byref(rect))
    cx = (rect.left + rect.right) // 2
    cy = (rect.top + rect.bottom) // 2
    prev = wt.POINT()
    GetCursorPos(ctypes.byref(prev))
    SetForegroundWindow(hwnd)
    time.sleep(0.15)
    SetCursorPos(cx, cy)
    time.sleep(0.1)
    mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.05)
    mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    time.sleep(0.2)
    SetCursorPos(prev.x, prev.y)


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
            if text in handled:
                return True
            handled.append(text)
            print("DIALOG:", text, "buttons:", buttons, flush=True)
            target = buttons[0][0]
            for h, t in buttons:
                if t.replace("&", "").strip().lower() in ("end", "ok"):
                    target = h
                    break
            click_control(target)
            return True

        EnumWindows(EnumProc(cb), 0)
        time.sleep(0.4)


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

try:
    excel.Run("SetupEmailCC")
    print("run ok", flush=True)
except Exception as e:
    print("COM ERROR:", str(e)[:150], flush=True)

time.sleep(1.5)

for attempt in range(12):
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
        print("introspect attempt", attempt, "failed:", str(e).replace("\n", " ")[:120], flush=True)
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
