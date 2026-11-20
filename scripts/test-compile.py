#!/usr/bin/env python3
"""test-compile.py

Compile-test harness for Pending_Pickup_Macro.bas.

Copies the template, injects a tiny probe macro, then runs it via COM.
Running ANY macro forces VBA to compile the ENTIRE project, so a compile
error anywhere blocks the probe.

Improvements over the previous version:
  - COM connectivity health check at start (DispatchEx + create + quit).
  - Fast in-process watcher thread (50 ms poll) instead of a slow
    subprocess watcher.
  - On any modal error dialog: captures dialog + VBE + code pane via
    PrintWindow (works even when another app is on top), kills Excel
    immediately to unblock the blocked excel.Run call, then OCRs the
    highlighted code line via scripts\\vba-error-ocr.py analyze.
  - Reports COMPILE OK or COMPILE FAILED with the dialog text, OCR'd
    highlighted line, and best source-line snap.

Safety:
  - A watchdog force-kills Excel if anything hangs.
  - Excel is always killed/closed at the end.
All output goes to log\\compile-test.log with per-line timestamps.
"""

import argparse
import ctypes
import ctypes.wintypes as wintypes
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
from datetime import datetime

import psutil
import pythoncom
import win32com.client
import win32gui
import win32process

parser = argparse.ArgumentParser(description="Compile-test Pending_Pickup_Macro.bas")
parser.add_argument("--template", default=None)
parser.add_argument("--timeout", type=int, default=25, help="Seconds to wait for the macro run.")
parser.add_argument("--no-ocr", action="store_true", help="Skip OCR of captured error UI.")
try:
    args = parser.parse_args()
except SystemExit:
    # Imported under pytest (sys.argv holds pytest args); fall back to defaults.
    args = parser.parse_args([])

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)

template = args.template or os.path.join(project_root, "Pending Pickup Template.xlsm")
test_copy = os.path.join(project_root, "_compile_test.xlsm")
marker = os.path.join(project_root, "probe_ran.txt")
log_dir = os.path.join(project_root, "log")
os.makedirs(log_dir, exist_ok=True)
watch_log = os.path.join(log_dir, "watcher.log")
test_log = os.path.join(log_dir, "compile-test.log")
ocr_script = os.path.join(script_dir, "vba-error-ocr.py")

user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32
kernel32 = ctypes.windll.kernel32

EnumWindows = user32.EnumWindows
EnumChildWindows = user32.EnumChildWindows
IsWindowVisible = user32.IsWindowVisible
GetWindowTextW = user32.GetWindowTextW
GetClassNameW = user32.GetClassNameW
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
PrintWindow = user32.PrintWindow
GetWindowDC = user32.GetWindowDC
ReleaseDC = user32.ReleaseDC
CreateCompatibleDC = gdi32.CreateCompatibleDC
CreateCompatibleBitmap = gdi32.CreateCompatibleBitmap
SelectObject = gdi32.SelectObject
GetDIBits = gdi32.GetDIBits
DeleteObject = gdi32.DeleteObject
DeleteDC = gdi32.DeleteDC


def log(msg):
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    with open(test_log, "a", encoding="utf-8") as f:
        f.write(line + "\n")


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
        EnumChildWindows(hwnd, ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)(cb), 0)
    except Exception:
        pass
    return texts


def _grab_printwindow(hwnd, flag):
    l, t, r, b = win32gui.GetWindowRect(hwnd)
    w, h = r - l, b - t
    if w <= 0 or h <= 0:
        return None
    hdc = GetWindowDC(hwnd)
    memdc = CreateCompatibleDC(hdc)
    bmp = CreateCompatibleBitmap(hdc, w, h)
    SelectObject(memdc, bmp)
    PrintWindow(hwnd, memdc, flag)

    class BIH(ctypes.Structure):
        _fields_ = [
            ("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG), ("biHeight", wintypes.LONG),
            ("biPlanes", wintypes.WORD), ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
            ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", wintypes.LONG), ("biYPelsPerMeter", wintypes.LONG),
            ("biClrUsed", wintypes.DWORD), ("biClrImportant", wintypes.DWORD),
        ]
    header = BIH()
    header.biSize = ctypes.sizeof(BIH)
    header.biWidth = w
    header.biHeight = -h
    header.biPlanes = 1
    header.biBitCount = 32
    buf = ctypes.create_string_buffer(w * h * 4)
    GetDIBits(memdc, bmp, 0, h, buf, ctypes.byref(header), 0)
    from PIL import Image
    img = Image.frombuffer("RGBA", (w, h), buf, "raw", "BGRA", 0, 1).convert("RGB")
    try:
        DeleteObject(bmp); DeleteDC(memdc); ReleaseDC(hwnd, hdc)
    except Exception:
        pass
    return img


def capture_window(hwnd, path):
    from PIL import ImageGrab
    for flag in (2, 0):
        img = _grab_printwindow(hwnd, flag)
        if img is not None:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            img.save(path)
            return path
    img = ImageGrab.grab(bbox=win32gui.GetWindowRect(hwnd), all_screens=True)
    if img is not None:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        img.save(path)
        return path
    return None


def excel_pids():
    pids = set()
    for p in psutil.process_iter(["name"]):
        try:
            if p.info["name"] and p.info["name"].lower() == "excel.exe":
                pids.add(p.pid)
        except Exception:
            pass
    return pids


def find_vbe_windows(pid):
    """Enumerate VBE main window + code pane child windows by class/title
    (no COM object access, safe from any thread)."""
    vbe = []
    code = []

    def enum_children(parent, cb):
        try:
            win32gui.EnumChildWindows(parent, ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)(cb), 0)
        except Exception:
            pass

    def cb(hwnd, _):
        try:
            if not IsWindowVisible(hwnd):
                return True
            wp = wintypes.DWORD()
            GetWindowThreadProcessId(hwnd, ctypes.byref(wp))
            if wp.value != pid:
                return True
            cls = win_class(hwnd)
            text = win_text(hwnd)
            if cls == "wndclass_desked_gsk":
                vbe.append((hwnd, text))
                # Enumerate children of this VBE main window for code panes.
                def child_cb(ch, _):
                    try:
                        ccls = win_class(ch)
                        ctext = win_text(ch)
                    except Exception:
                        return True
                    if ccls == "VbaWindow" and ctext.endswith("(Code)"):
                        code.append((ch, ctext))
                        return False  # stop after first code pane
                    return True
                enum_children(hwnd, child_cb)
        except Exception:
            pass
        return True

    try:
        EnumWindows(ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)(cb), 0)
    except Exception:
        pass
    return vbe, code


def find_dialogs(pid):
    found = []
    def cb(hwnd, _):
        try:
            if not IsWindowVisible(hwnd):
                return True
            wp = wintypes.DWORD()
            GetWindowThreadProcessId(hwnd, ctypes.byref(wp))
            if wp.value != pid:
                return True
            if win_class(hwnd) != "#32770":
                return True
            texts = dialog_child_texts(hwnd)
            found.append((hwnd, win_text(hwnd), texts))
        except Exception:
            pass
        return True
    try:
        EnumWindows(ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)(cb), 0)
    except Exception:
        pass
    return found


def kill_excel():
    for p in psutil.process_iter(["name", "pid"]):
        try:
            if (p.info["name"] or "").lower() == "excel.exe":
                p.kill()
        except Exception:
            pass


def ocr_image(image_path, module, sources, outdir, label):
    if not os.path.isfile(image_path):
        return []
    subprocess.run(
        [sys.executable, ocr_script, "analyze", image_path, "--module", module or "",
         "--source", *sources, "--out", outdir],
        capture_output=True, text=True, timeout=90,
    )
    results = []
    report = os.path.join(outdir, "report.json")
    if os.path.isfile(report):
        try:
            with open(report, "r", encoding="utf-8") as f:
                data = json.load(f)
            for finding in data.get("findings", []):
                for band in finding.get("bands", []):
                    best = band.get("best")
                    if best and best.get("matches"):
                        top = best["matches"][0]
                        results.append(
                            f"match {top['score']:.3f}: {top['file']}:{top['line']} — {top['source']} "
                            f"(OCR: {best['text']})"
                        )
        except Exception:
            pass
    return results


def cmd_health():
    """COM connectivity health check."""
    log("=== COM health check ===")
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Add()
        log("COM health: DispatchEx + Workbooks.Add OK")
        wb.Close(False)
        log("COM health: Workbook add/close OK")
        return True
    except Exception as e:
        log(f"COM health FAILED: {e}")
        return False
    finally:
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
        kill_excel()


def cmd_collect(template_path):
    """Full compile test with OCR capture on failure."""
    log(f"=== Compile test started (template: {template_path}) ===")

    if not cmd_health():
        log("RESULT: COMPILE TEST ABORTED - COM health check failed.")
        return 1

    kill_excel()
    time.sleep(1)

    # Fresh test copy
    if os.path.exists(test_copy):
        os.remove(test_copy)
    shutil.copy2(template_path, test_copy)
    if os.path.exists(marker):
        os.remove(marker)

    captured = {}
    stop = threading.Event()
    deadline = time.time() + max(5, args.timeout)

    excel = None
    wb = None
    pid = None
    run_error = None

    try:
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        pid = win32process.GetWindowThreadProcessId(excel.Hwnd)[1]
        wb = excel.Workbooks.Open(test_copy)
        vbp = wb.VBProject
        mod = vbp.VBComponents.Add(1)
        mod.Name = "zzCompileProbe"
        probe_code = (
            "Option Explicit\r\n"
            "Public Sub Probe()\r\n"
            f'    Open "{marker}" For Output As #1\r\n'
            '    Print #1, "ran"\r\n'
            "    Close #1\r\n"
            "End Sub"
        )
        mod.CodeModule.AddFromString(probe_code)
        log("Injecting probe module zzCompileProbe (forces full-project compile)...")

        # Show VBE so the code pane window exists for capture
        try:
            excel.VBE.MainWindow.Visible = True
            mod.CodeModule.CodePane.Show()
        except Exception as exc:
            log(f"VBE show warning: {exc}")

        def watcher():
            try:
                while not stop.is_set() and time.time() < deadline:
                    dialogs = find_dialogs(pid)
                    if dialogs:
                        hwnd, title, texts = dialogs[0]
                        captured["dialog_title"] = title
                        captured["dialog_lines"] = texts
                        captured["dialog_image"] = capture_window(hwnd, os.path.join(log_dir, "compile-dialog.png"))
                        # Capture VBE + code pane via window enumeration (no COM
                        # access from this thread).
                        vbe_wins, code_wins = find_vbe_windows(pid)
                        if vbe_wins:
                            vbe_hwnd, vbe_title = vbe_wins[0]
                            captured["vbe_image"] = capture_window(vbe_hwnd, os.path.join(log_dir, "compile-vbe.png"))
                        if code_wins:
                            c_hwnd, c_caption = code_wins[0]
                            captured["code_caption"] = c_caption
                            captured["code_image"] = capture_window(c_hwnd, os.path.join(log_dir, "compile-code.png"))
                        kill_excel()
                        stop.set()
                        return
                    time.sleep(0.05)
            finally:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

        watcher_thread = threading.Thread(target=watcher, daemon=True)
        watcher_thread.start()

        try:
            excel.Run("Probe")
        except Exception as e:
            run_error = str(e)
            log(f"Run call raised: {run_error}")
        watcher_thread.join(timeout=args.timeout + 5)
    finally:
        stop.set()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        kill_excel()

    # Wait briefly for watcher to finish capturing
    time.sleep(1)

    compile_ok = os.path.exists(marker)
    if compile_ok:
        log("RESULT: COMPILE OK - probe executed, project compiled cleanly.")
        rc = 0
    else:
        log("RESULT: COMPILE FAILED - probe did not execute.")
        rc = 1
        if captured.get("dialog_lines"):
            log("Captured dialog text:")
            for l in captured["dialog_lines"]:
                log(f"  {l}")
        if run_error:
            log(f"Run error: {run_error}")
        # OCR the captured code pane
        if not args.no_ocr and captured.get("code_image") and os.path.isfile(captured["code_image"]):
            outdir = os.path.join(log_dir, "compile-ocr")
            os.makedirs(outdir, exist_ok=True)
            module = captured.get("code_caption")
            # extract module name from caption like "Probe.xlsm - zzCompileProbe (Code)"
            import re
            m = re.search(r" - (\w+) \(Code\)", module or "")
            module_name = m.group(1) if m else None
            sources = [os.path.join(project_root, "vba", f) for f in os.listdir(os.path.join(project_root, "vba"))] if os.path.isdir(os.path.join(project_root, "vba")) else []
            results = ocr_image(captured["code_image"], module_name, sources, outdir, "compile")
            log("OCR results:")
            for r in results:
                log(f"  {r}")
        if not captured.get("dialog_lines") and not captured.get("code_image"):
            log("No error dialog captured.")

    # Cleanup
    try:
        if wb:
            wb.Close(False)
    except Exception:
        pass
    try:
        if excel:
            excel.Quit()
    except Exception:
        pass
    kill_excel()
    try:
        os.remove(test_copy)
    except OSError:
        pass
    try:
        os.remove(marker)
    except OSError:
        pass
    log("=== Compile test finished ===")
    return rc


def main():
    return cmd_collect(template)


if __name__ == "__main__":
    sys.exit(main())
