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
test_copy = os.path.join(temp_dir, "shot_template.xlsm")
shots = os.path.join(temp_dir, "shots")
os.makedirs(shots, exist_ok=True)

for p in psutil.process_iter(["name"]):
    if p.info["name"] and p.info["name"].lower() == "excel.exe":
        p.kill()
if os.path.exists(test_copy):
    os.remove(test_copy)
shutil.copy2(template, test_copy)

stop = threading.Event()
seen_dialog = threading.Event()
captured = []


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
    PrintWindow(hwnd, memdc, 2)

    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [
            ("biSize", wt.DWORD), ("biWidth", wt.LONG), ("biHeight", wt.LONG),
            ("biPlanes", wt.WORD), ("biBitCount", wt.WORD), ("biCompression", wt.DWORD),
            ("biSizeImage", wt.DWORD), ("biXPelsPerMeter", wt.LONG), ("biYPelsPerMeter", wt.LONG),
            ("biClrUsed", wt.DWORD), ("biClrImportant", wt.DWORD),
        ]

    bi = BITMAPINFOHEADER()
    bi.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    bi.biWidth = w
    bi.biHeight = -h
    bi.biPlanes = 1
    bi.biBitCount = 32
    buf = ctypes.create_string_buffer(w * h * 4)
    gdi32.GetDIBits(memdc, bmp, 0, h, buf, ctypes.byref(bi), 0)
    img = Image.frombuffer("RGBA", (w, h), buf, "raw", "BGRA", 0, 1)
    img.convert("RGB").save(path)
    gdi32.DeleteObject(bmp)
    gdi32.DeleteDC(memdc)
    user32.ReleaseDC(hwnd, hdc)
    return True


def watcher(pid):
    idx = 0
    while not stop.is_set():
        def cb(hwnd, _):
            nonlocal idx
            p = wt.DWORD()
            GetWindowThreadProcessId(hwnd, ctypes.byref(p))
            if p.value != pid or not IsWindowVisible(hwnd):
                return True
            cls = win_class(hwnd)
            text = win_text(hwnd)
            if cls in ("#32770", "wndclass_desked_gsk", "XLMAIN"):
                if cls == "#32770":
                    seen_dialog.set()
                if text not in [c[0] for c in captured]:
                    captured.append((text, idx))
                    idx += 1
                    fn = os.path.join(shots, f"{idx:02d}_{cls.strip('#').replace(' ', '_')}.png")
                    try:
                        capture(hwnd, fn)
                        print("captured", fn, "|", cls, "|", text, flush=True)
                    except Exception as e:
                        print("capture failed", e, flush=True)
            return True

        EnumWindows(EnumProc(cb), 0)
        time.sleep(0.5)
        if seen_dialog.is_set():
            time.sleep(1.0)
            return


captured = []
pythoncom.CoInitialize()
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(test_copy)
pid = wt.DWORD()
GetWindowThreadProcessId(excel.Hwnd, ctypes.byref(pid))
pid = pid.value
print("Excel pid:", pid, flush=True)
th = threading.Thread(target=watcher, args=(pid,), daemon=True)
th.start()

run_thread = threading.Thread(target=lambda: None)
try:
    excel.Run("SetupEmailCC")
    print("run ok", flush=True)
except Exception as e:
    print("COM ERROR:", e, flush=True)
stop.set()
time.sleep(0.5)
try:
    excel.Quit()
except Exception:
    pass
pythoncom.CoUninitialize()
for p in psutil.process_iter(["name"]):
    if p.info["name"] and p.info["name"].lower() == "excel.exe":
        p.kill()
print("done")
