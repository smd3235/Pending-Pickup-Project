import ctypes
import ctypes.wintypes as wt
import os
import shutil
import threading
import time

import psutil
import pythoncom
import win32com.client
from PIL import ImageGrab

user32 = ctypes.windll.user32
EnumWindows = user32.EnumWindows
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
GetWindowTextW = user32.GetWindowTextW
GetClassNameW = user32.GetClassNameW
IsWindowVisible = user32.IsWindowVisible
SetForegroundWindow = user32.SetForegroundWindow
ShowWindow = user32.ShowWindow
SW_RESTORE = 9

EnumProc = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)

root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
template = os.path.join(root, "Pending Pickup Template.xlsm")
temp_dir = os.path.join(root, "scripts", "temp")
test_copy = os.path.join(temp_dir, "fullscreen_hunt.xlsm")
out_png = os.path.join(temp_dir, "error_screen.png")

for p in psutil.process_iter(["name"]):
    if p.info["name"] and p.info["name"].lower() == "excel.exe":
        p.kill()
if os.path.exists(test_copy):
    os.remove(test_copy)
shutil.copy2(template, test_copy)

stop = threading.Event()
captured = threading.Event()


def win_text(hwnd):
    n = user32.GetWindowTextLengthW(hwnd)
    b = ctypes.create_unicode_buffer(n + 1)
    user32.GetWindowTextW(hwnd, b, n + 1)
    return b.value


def win_class(hwnd):
    b = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, b, 256)
    return b.value


def find_windows(pid):
    found = []

    def cb(hwnd, _):
        p = wt.DWORD()
        GetWindowThreadProcessId(hwnd, ctypes.byref(p))
        if p.value == pid and IsWindowVisible(hwnd):
            found.append((hwnd, win_class(hwnd), win_text(hwnd)))
        return True

    EnumWindows(EnumProc(cb), 0)
    return found


def watcher(pid):
    while not stop.is_set():
        wins = find_windows(pid)
        if any(cls == "#32770" for _, cls, _ in wins):
            vbe = [hwnd for hwnd, cls, _ in wins if cls == "wndclass_desked_gsk"]
            if vbe:
                ShowWindow(vbe[0], SW_RESTORE)
                SetForegroundWindow(vbe[0])
            time.sleep(0.8)
            img = ImageGrab.grab(all_screens=True)
            img.save(out_png)
            print("saved screenshot", out_png, img.size, flush=True)
            captured.set()
            time.sleep(0.5)
            for p in psutil.process_iter(["name"]):
                if p.info["name"] and p.info["name"].lower() == "excel.exe":
                    p.kill()
            return
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

try:
    excel.VBE.MainWindow.Visible = True
    vbe = excel.VBE
    vbe.ActiveCodePane = wb.VBProject.VBComponents("modDialog").CodeModule.CodePane
    vbe.ActiveCodePane.Show()
    print("VBE shown", flush=True)
except Exception as e:
    print("VBE setup failed:", e, flush=True)

threading.Thread(target=watcher, args=(pid,), daemon=True).start()

try:
    excel.Run("SetupEmailCC")
    print("run ok", flush=True)
except Exception as e:
    print("COM ERROR:", str(e)[:200], flush=True)

captured.wait(timeout=20)
stop.set()
pythoncom.CoUninitialize()
for p in psutil.process_iter(["name"]):
    if p.info["name"] and p.info["name"].lower() == "excel.exe":
        p.kill()
print("done", flush=True)
