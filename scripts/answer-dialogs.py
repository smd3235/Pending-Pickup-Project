# answer-dialogs.py
# Answers the report macro's modal dialogs so a report run can complete unattended.
# Handles:
#  - Excel FileDialog picker (class bosa_sdm_MsoFiles): type path into the File name edit + Open
#  - #32770 dialogs: CURRENT/PRIOR pickers by title, other dialogs -> Enter (logged by watcher)
# Log: log\answer-dialogs.log (per-line timestamps).

import argparse
import ctypes
import ctypes.wintypes as wintypes
import os
import re
import subprocess
import sys
import threading
import time
from datetime import datetime

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

EnumWindows = user32.EnumWindows
EnumChildWindows = user32.EnumChildWindows
IsWindowVisible = user32.IsWindowVisible
GetWindowTextW = user32.GetWindowTextW
GetClassNameW = user32.GetClassNameW
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
SetForegroundWindow = user32.SetForegroundWindow
SendMessageW = user32.SendMessageW
PostMessageW = user32.PostMessageW

EnumProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

WM_SETTEXT = 0x000C
WM_COMMAND = 0x0111
WM_KEYDOWN = 0x0100
IDOK = 1
VK_ESCAPE = 0x1B


def excel_pids():
    try:
        import psutil
        pids = set()
        for p in psutil.process_iter(["name"]):
            try:
                if p.info["name"] and p.info["name"].lower() == "excel.exe":
                    pids.add(p.pid)
            except Exception:
                pass
        return pids
    except ImportError:
        return set()


def win_text(hwnd):
    buf = ctypes.create_unicode_buffer(1024)
    GetWindowTextW(hwnd, buf, 1024)
    return buf.value


def win_class(hwnd):
    buf = ctypes.create_unicode_buffer(256)
    GetClassNameW(hwnd, buf, 256)
    return buf.value


def dialog_child_texts(hwnd):
    texts = []

    def cb(child, _):
        try:
            cls = win_class(child)
            text = win_text(child).strip()
        except Exception:
            return True
        if text and cls in ("Static", "Button", "Edit", "SysLink"):
            if text not in texts:
                texts.append(text)
        return True

    try:
        EnumChildWindows(hwnd, EnumProc(cb), 0)
    except Exception:
        pass
    return texts


def is_vba_error_dialog(hwnd):
    title = win_text(hwnd)
    texts = dialog_child_texts(hwnd)
    if "error" in title.lower():
        return True
    return any("error" in t.lower() for t in texts)


def parse_error_number(texts):
    for t in texts:
        m = re.search(r"error\s*(?:\(\s*'?(\d+)'?\s*\)|:\s*(\d+))", t, re.IGNORECASE)
        if m:
            return m.group(1) or m.group(2)
    return None


_project_root = os.path.dirname(os.path.abspath(__file__))
_parent_root = os.path.dirname(_project_root)
_vba_ocr = os.path.join(_project_root, "scripts", "vba-error-ocr.py")
_ocr_lock = threading.Lock()
_ocr_in_flight = False


def capture_vba_error_live(hwnd):
    """Snapshot the live VBA UI behind any VBA error dialog (non-blocking)."""
    global _ocr_in_flight
    with _ocr_lock:
        if _ocr_in_flight:
            return
        _ocr_in_flight = True

    def worker():
        try:
            pid = wintypes.DWORD()
            GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            subprocess.run(
                [
                    sys.executable,
                    _vba_ocr,
                    "ocr-live",
                    "--pid",
                    str(pid.value),
                    "--out",
                    os.path.join(_parent_root, "log", "ocr", f"live-{int(time.time())}"),
                ],
                capture_output=True,
                text=True,
                timeout=45,
            )
        except Exception as exc:
            try:
                with open(os.path.join(_parent_root, "log", "answer-dialogs.log"), "a", encoding="utf-8") as f:
                    f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | live OCR capture failed: {exc}\n")
            except Exception:
                pass
        finally:
            with _ocr_lock:
                _ocr_in_flight = False

    threading.Thread(target=worker, daemon=True).start()


def find_child(hwnd, cls_name):
    found = [None]

    def cb(child, _):
        if win_class(child) == cls_name:
            found[0] = child
            return False
        return True

    EnumChildWindows(hwnd, EnumProc(cb), 0)
    return found[0]


def excel_dialogs():
    pids = excel_pids()
    if not pids:
        return []
    result = []

    def cb(hwnd, _):
        if not IsWindowVisible(hwnd):
            return True
        pid = wintypes.DWORD()
        GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        if pid.value not in pids:
            return True
        result.append(hwnd)
        return True

    EnumWindows(EnumProc(cb), 0)
    return result


def send_keys(text):
    """Send keystrokes via SendInput."""
    inputs = []
    VkKeyScanW = ctypes.windll.user32.VkKeyScanW
    MapVirtualKeyW = ctypes.windll.user32.MapVirtualKeyW

    class KEYBDINPUT(ctypes.Structure):
        _fields_ = [
            ("wVk", wintypes.WORD),
            ("wScan", wintypes.WORD),
            ("dwFlags", wintypes.DWORD),
            ("time", wintypes.DWORD),
            ("dwExtraInfo", wintypes.ULONG_PTR),
        ]

    class INPUT(ctypes.Structure):
        _fields_ = [
            ("type", wintypes.DWORD),
            ("ki", KEYBDINPUT),
        ]

    INPUT_KEYBOARD = 1
    KEYEVENTF_KEYUP = 0x0002

    def key_down(vk):
        inputs.append(INPUT(1, KEYBDINPUT(vk, 0, 0, 0, 0)))

    def key_up(vk):
        inputs.append(INPUT(1, KEYBDINPUT(vk, 0, KEYEVENTF_KEYUP, 0, 0)))

    for ch in text:
        vk = VkKeyScanW(ch)
        key_down(vk & 0xFF)
        key_up(vk & 0xFF)

    if inputs:
        n = len(inputs)
        arr = (INPUT * n)(*inputs)
        user32.SendInput(n, arr, ctypes.sizeof(INPUT))


def key_down_once(vk):
    class KEYBDINPUT(ctypes.Structure):
        _fields_ = [
            ("wVk", wintypes.WORD),
            ("wScan", wintypes.WORD),
            ("dwFlags", wintypes.DWORD),
            ("time", wintypes.DWORD),
            ("dwExtraInfo", wintypes.ULONG_PTR),
        ]

    class INPUT(ctypes.Structure):
        _fields_ = [
            ("type", wintypes.DWORD),
            ("ki", KEYBDINPUT),
        ]

    INPUT_KEYBOARD = 1
    KEYEVENTF_KEYUP = 0x0002

    ki_down = KEYBDINPUT(vk, 0, 0, 0, 0)
    ki_up = KEYBDINPUT(vk, 0, KEYEVENTF_KEYUP, 0, 0)
    arr = (INPUT * 2)(INPUT(1, ki_down), INPUT(1, ki_up))
    user32.SendInput(2, arr, ctypes.sizeof(INPUT))


project_root = os.path.dirname(os.path.abspath(__file__))
parent_root = os.path.dirname(project_root)
log_dir = os.path.join(parent_root, "log")
os.makedirs(log_dir, exist_ok=True)

parser = argparse.ArgumentParser(description="Answer Excel modal dialogs unattended")
parser.add_argument("--source", default=os.path.join(parent_root, "AllOpenTreatmentPlansResults618.xlsx"))
parser.add_argument("--prior", default="")
parser.add_argument("--log", default=os.path.join(log_dir, "answer-dialogs.log"))
args = parser.parse_args()

source_path = args.source
prior_path = args.prior
log_path = args.log


def log(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(line + "\n")


log(f"answerer started (pid {os.getpid()})")
answered = {}
deadline = time.time() + 10 * 60

try:
    while time.time() < deadline:
        time.sleep(0.4)
        live = {}
        for h in excel_dialogs():
            live[str(h)] = h
        for k in list(answered.keys()):
            if k not in live:
                del answered[k]
        for k in live:
            if k in answered:
                continue
            h = live[k]
            title = win_text(h)
            cls = win_class(h)
            if cls == "bosa_sdm_MsoFiles":
                time.sleep(0.5)
                SetForegroundWindow(h)
                time.sleep(0.2)
                edit = find_child(h, "Edit")
                if edit is not None:
                    SendMessageW(edit, WM_SETTEXT, 0, source_path)
                SendMessageW(h, WM_COMMAND, IDOK, 0)
                answered[k] = 1
                log(f"answered picker (title={title}) -> {source_path}")
            elif "Select the CURRENT" in title:
                time.sleep(0.5)
                SetForegroundWindow(h)
                send_keys("^a")
                send_keys(source_path)
                key_down_once(0x0D)  # VK_RETURN
                answered[k] = 1
                log(f"answered CURRENT picker -> {source_path}")
            elif "Select the PRIOR" in title:
                if prior_path:
                    time.sleep(0.4)
                    SetForegroundWindow(h)
                    send_keys("^a")
                    send_keys(prior_path)
                    key_down_once(0x0D)
                    log(f"answered PRIOR picker -> {prior_path}")
                else:
                    time.sleep(0.4)
                    SetForegroundWindow(h)
                    key_down_once(VK_ESCAPE)
                    log("cancelled PRIOR picker (no prior report available)")
                answered[k] = 1
            elif "Choose where" in title:
                time.sleep(0.4)
                SetForegroundWindow(h)
                key_down_once(0x0D)
                answered[k] = 1
                log("accepted default output folder")
            else:
                time.sleep(0.3)
                texts = dialog_child_texts(h)
                err_no = parse_error_number(texts)
                if is_vba_error_dialog(h):
                    capture_vba_error_live(h)
                    log(f"captured live VBA UI for VBA error {err_no or '?'} (title={title})")
                SetForegroundWindow(h)
                key_down_once(0x0D)
                answered[k] = 1
                log(f"dismissed dialog (class={cls}): {title}" + (f" | error={err_no}" if err_no else ""))
finally:
    log("answerer stopped")
