import ctypes
import ctypes.wintypes as wt
import os
import shutil
import threading
import time

import psutil
import pythoncom
import win32com.client
from PIL import Image

user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32
EnumWindows = user32.EnumWindows
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
GetWindowTextW = user32.GetWindowTextW
GetClassNameW = user32.GetClassNameW
GetWindowRect = user32.GetWindowRect
IsWindowVisible = user32.IsWindowVisible
PrintWindow = user32.PrintWindow

EnumProc = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)

root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
template = os.path.join(root, "Pending Pickup Template.xlsm")
temp_dir = os.path.join(root, "scripts", "temp")
test_copy = os.path.join(temp_dir, "vbe_hunt.xlsm")
shots = os.path.join(temp_dir, "shots2")
os.makedirs(shots, exist_ok=True)
for f in os.listdir(shots):
    os.remove(os.path.join(shots, f))

for p in psutil.process_iter(["name"]):
    if p.info["name"] and p.info["name"].lower() == "excel.exe":
        p.kill()
if os.path.exists(test_copy):
    os.remove(test_copy)
shutil.copy2(template, test_copy)

stop = threading.Event()
done = threading.Event()


def win_text(hwnd):
    n = user32.GetWindowTextLengthW(hwnd)
    b = ctypes.create_unicode_buffer(n + 1)
    user32.GetWindowTextW(hwnd, b, n + 1)
    return b.value


def win_class(hwnd):
    b = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, b, 256)
    return b.value


def capture(hwnd, path):
    rect = wt.RECT()
    GetWindowRect(hwnd, ctypes.byref(rect))
    w = rect.right - rect.left
    h = rect.bottom - rect.top
    if w <= 0 or h <= 0:
        return False
    hdc = user32.GetWindowDC(hwnd)
    memdc = gdi32.CreateCompatibleDC(hdc)
    bmp = gdi32.CreateCompatibleBitmap(hdc, w, h)
    gdi32.SelectObject(memdc, bmp)
    PrintWindow(hwnd, memdc, 3)

    class BIH(ctypes.Structure):
        _fields_ = [("biSize", wt.DWORD), ("biWidth", wt.LONG), ("biHeight", wt.LONG),
                    ("biPlanes", wt.WORD), ("biBitCount", wt.WORD), ("biCompression", wt.DWORD),
                    ("biSizeImage", wt.DWORD), ("biXPelsPerMeter", wt.LONG), ("biYPelsPerMeter", wt.LONG),
                    ("biClrUsed", wt.DWORD), ("biClrImportant", wt.DWORD)]

    bi = BIH()
    bi.biSize = ctypes.sizeof(BIH)
    bi.biWidth = w
    bi.biHeight = -h
    bi.biPlanes = 1
    bi.biBitCount = 32
    buf = ctypes.create_string_buffer(w * h * 4)
    gdi32.GetDIBits(memdc, bmp, 0, h, buf, ctypes.byref(bi), 0)
    Image.frombuffer("RGBA", (w, h), buf, "raw", "BGRA", 0, 1).convert("RGB").save(path)
    gdi32.DeleteObject(bmp)
    gdi32.DeleteDC(memdc)
    user32.ReleaseDC(hwnd, hdc)
    return True


def watcher(pid):
    idx = 0
    while not stop.is_set():
        wins = []

        def cb(hwnd, _):
            p = wt.DWORD()
            GetWindowThreadProcessId(hwnd, ctypes.byref(p))
            if p.value == pid and IsWindowVisible(hwnd):
                wins.append((hwnd, win_class(hwnd), win_text(hwnd)))
            return True

        EnumWindows(EnumProc(cb), 0)
        has_dialog = any(cls == "#32770" for _, cls, _ in wins)
        if has_dialog:
            time.sleep(1.0)
            for hwnd, cls, text in wins:
                idx += 1
                safe = cls.strip("#").replace(" ", "_")
                fn = os.path.join(shots, f"{idx:02d}_{safe}.png")
                try:
                    if capture(hwnd, fn):
                        print("captured", fn, "|", cls, "|", text, flush=True)
                except Exception as e:
                    print("capture fail", e, flush=True)
            done.set()
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
    print("COM ERROR:", str(e)[:150], flush=True)

done.wait(timeout=20)
stop.set()
pythoncom.CoUninitialize()
for p in psutil.process_iter(["name"]):
    if p.info["name"] and p.info["name"].lower() == "excel.exe":
        p.kill()
print("done", flush=True)
