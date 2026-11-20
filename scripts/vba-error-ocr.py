"""Capture and OCR VBA UI errors (dialog text + highlighted code line).

Usage:
  python scripts\\vba-error-ocr.py collect --macro SetupEmailCC --module modDialog
    Launch a probe workbook, run the macro, capture the error UI on the first
    modal compile dialog, kill Excel to unblock, then OCR + snap to source.
  python scripts\\vba-error-ocr.py ocr-live --pid 1234
    Attach to a *running* Excel and capture its live VBA UI (dialog + VBE +
    code pane + desktop + all VBA windows), OCR the highlighted code band,
    snap to vba/*.bas sources. Does NOT kill Excel.
  python scripts\\vba-error-ocr.py ocr-all
    Batch-OCR every screenshot already under log\\ocr\\ into one combined report.
  python scripts\\vba-error-ocr.py analyze log\\ocr\\<ts>\\vbe-code.png --module modX

Artifacts are written to log\\ocr\\<timestamp>\\ and summarized in log\\vba-error-ocr.log.
"""

import argparse
import ctypes
import ctypes.wintypes as wt
import difflib
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
from PIL import Image, ImageFilter, ImageGrab, ImageOps

from vba_ocr_helpers import (
    detect_bands,
    flatten_selection,
    load_sources,
    normalize,
    preprocess_variants,
    snap_to_source,
    discover_sources,
)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMP_DIR = os.path.join(ROOT, "scripts", "temp")
LOG_DIR = os.path.join(ROOT, "log")
OCR_ROOT = os.path.join(LOG_DIR, "ocr")
REPORT_LOG = os.path.join(LOG_DIR, "vba-error-ocr.log")
OCR_PS1 = os.path.join(ROOT, "scripts", "ocr-winrt.ps1")
DEFAULT_WORKBOOK = os.path.join(ROOT, "Pending Pickup Template.xlsm")

user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32


def enable_dpi_awareness():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            user32.SetProcessDPIAware()
        except Exception:
            pass


def kill_excel():
    killed = []
    for proc in psutil.process_iter(["name", "pid"]):
        if (proc.info["name"] or "").lower() == "excel.exe":
            try:
                proc.kill()
                killed.append(proc.info["pid"])
            except Exception:
                pass
    return killed


def window_rect(hwnd):
    try:
        return win32gui.GetWindowRect(hwnd)
    except Exception:
        return None


def _grab_printwindow(hwnd, flag):
    rect = window_rect(hwnd)
    if not rect:
        return None
    left, top, right, bottom = rect
    width, height = right - left, bottom - top
    if width <= 0 or height <= 0:
        return None
    hdc = None
    memdc = None
    bmp = None
    try:
        hdc = user32.GetWindowDC(hwnd)
        memdc = gdi32.CreateCompatibleDC(hdc)
        bmp = gdi32.CreateCompatibleBitmap(hdc, width, height)
        gdi32.SelectObject(memdc, bmp)
        user32.PrintWindow(hwnd, memdc, flag)

        class BITMAPINFOHEADER(ctypes.Structure):
            _fields_ = [
                ("biSize", wt.DWORD),
                ("biWidth", wt.LONG),
                ("biHeight", wt.LONG),
                ("biPlanes", wt.WORD),
                ("biBitCount", wt.WORD),
                ("biCompression", wt.DWORD),
                ("biSizeImage", wt.DWORD),
                ("biXPelsPerMeter", wt.LONG),
                ("biYPelsPerMeter", wt.LONG),
                ("biClrUsed", wt.DWORD),
                ("biClrImportant", wt.DWORD),
            ]

        header = BITMAPINFOHEADER()
        header.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        header.biWidth = width
        header.biHeight = -height
        header.biPlanes = 1
        header.biBitCount = 32
        buffer = ctypes.create_string_buffer(width * height * 4)
        gdi32.GetDIBits(memdc, bmp, 0, height, buffer, ctypes.byref(header), 0)
        return Image.frombuffer("RGBA", (width, height), buffer, "raw", "BGRA", 0, 1).convert("RGB")
    except Exception:
        return None
    finally:
        try:
            if bmp:
                gdi32.DeleteObject(bmp)
        except Exception:
            pass
        try:
            if memdc:
                gdi32.DeleteDC(memdc)
        except Exception:
            pass
        try:
            if hdc:
                user32.ReleaseDC(hwnd, hdc)
        except Exception:
            pass


def _grab_screen(hwnd):
    rect = window_rect(hwnd)
    if not rect:
        return None
    try:
        return ImageGrab.grab(bbox=rect, all_screens=True)
    except Exception:
        return None


def _grab_desktop():
    try:
        return ImageGrab.grab(all_screens=True)
    except Exception:
        return None


def _is_blank(img):
    if img is None:
        return True
    gray = img.convert("L")
    low, high = gray.getextrema()
    return (high - low) < 5


def capture_window(hwnd, path):
    attempts = [
        lambda: _grab_printwindow(hwnd, 2),
        lambda: _grab_printwindow(hwnd, 0),
        lambda: _grab_screen(hwnd),
    ]
    for attempt in attempts:
        img = attempt()
        if img is not None and not _is_blank(img):
            os.makedirs(os.path.dirname(path), exist_ok=True)
            img.save(path)
            return path
    img = attempts[-1]()
    if img is not None:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        img.save(path)
        return path
    return None


def capture_desktop(path):
    img = _grab_desktop()
    if img is not None:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        img.save(path)
        return path
    return None


def dialog_text(hwnd):
    texts = []

    def callback(child, _):
        try:
            cls = win32gui.GetClassName(child)
            text = win32gui.GetWindowText(child).strip()
        except Exception:
            return True
        if text and cls in ("Static", "Button", "Edit", "SysLink", "DirectUIHWND"):
            if text not in texts:
                texts.append(text)
        return True

    try:
        win32gui.EnumChildWindows(hwnd, callback, None)
    except Exception:
        pass
    if not texts:
        title = win32gui.GetWindowText(hwnd).strip()
        if title:
            texts.append(title)
    return texts


def find_excel_windows(pid):
    found = {"dialogs": [], "vbe": [], "excel": [], "all": []}
    if not pid:
        return found

    def callback(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        try:
            _, win_pid = win32process.GetWindowThreadProcessId(hwnd)
            if win_pid != pid:
                return True
            cls = win32gui.GetClassName(hwnd)
            text = win32gui.GetWindowText(hwnd)
        except Exception:
            return True
        found["all"].append((hwnd, cls, text))
        if cls == "#32770":
            found["dialogs"].append((hwnd, text, dialog_text(hwnd)))
        elif cls == "wndclass_desked_gsk":
            found["vbe"].append((hwnd, text))
        elif cls == "XLMAIN":
            found["excel"].append((hwnd, text))
        return True

    try:
        win32gui.EnumWindows(callback, None)
    except Exception:
        pass
    return found


def find_code_windows(vbe_hwnd):
    code_windows = []

    def callback(hwnd, _):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            cls = win32gui.GetClassName(hwnd)
            text = win32gui.GetWindowText(hwnd)
        except Exception:
            return True
        if cls == "VbaWindow" and text.endswith("(Code)"):
            code_windows.append((hwnd, text))
        return True

    try:
        win32gui.EnumChildWindows(vbe_hwnd, callback, None)
    except Exception:
        pass
    return code_windows


def module_from_caption(caption):
    match = re.search(r" - (\w+) \(Code\)\s*$", caption or "")
    return match.group(1) if match else None

def ocr_words(path):
    proc = subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            OCR_PS1,
            "-Path",
            path,
        ],
        capture_output=True,
        text=True,
        timeout=30,
    )
    lines = []
    words = []
    for raw in proc.stdout.splitlines():
        if "\t" not in raw:
            continue
        kind, box, text = raw.split("\t", 2)
        try:
            x, y, w, h = (int(v) for v in box.split(","))
        except ValueError:
            continue
        item = {"x": x, "y": y, "w": w, "h": h, "text": text}
        if kind == "LINE":
            lines.append(item)
        elif kind == "WORD":
            words.append(item)
    return {"lines": lines, "words": words, "stderr": proc.stderr.strip(), "returncode": proc.returncode}


def analyze_image(image_path, module, source_paths, outdir, label):
    with Image.open(image_path) as opened:
        img = opened.convert("RGB")
    source_entries = load_sources(source_paths)
    findings = {
        "label": label,
        "image": os.path.relpath(image_path, ROOT),
        "size": list(img.size),
        "bands": [],
        "full_pane": [],
        "module": module,
    }
    bands = detect_bands(img)
    for index, band in enumerate(bands, start=1):
        y0 = max(0, band["y0"] - 5)
        y1 = min(img.size[1], band["y1"] + 5)
        crop_specs = [
            ("token", (max(0, band["x0"] - 16), y0, min(img.size[0], band["x1"] + 32), y1)),
            ("line", (0, y0, img.size[0], y1)),
        ]
        best = None
        best_rank = (-1.0, -1)
        candidates = []
        for crop_name, crop_box in crop_specs:
            crop = img.crop(crop_box)
            for scale in (3, 4, 5):
                variants = preprocess_variants(crop, scale)
                if crop_name == "token":
                    flat = flatten_selection(crop, band["color"])
                    flat = flat.resize(
                        (flat.size[0] * scale, flat.size[1] * scale), Image.LANCZOS
                    )
                    variants["flattened"] = ImageOps.autocontrast(flat)
                for variant_name, variant in variants.items():
                    variant_path = os.path.join(
                        outdir, f"{label}-band{index}-{crop_name}-{scale}x-{variant_name}.png"
                    )
                    variant.save(variant_path)
                    ocr = ocr_words(variant_path)
                    text = " ".join(line["text"] for line in ocr["lines"]).strip()
                    if not text:
                        continue
                    matches = snap_to_source(text, source_entries)
                    score = matches[0]["score"] if matches else 0.0
                    candidates.append(
                        {
                            "crop": crop_name,
                            "scale": scale,
                            "variant": variant_name,
                            "image": os.path.relpath(variant_path, ROOT),
                            "text": text,
                            "matches": matches,
                            "score": score,
                        }
                    )
                    rank = (score, len(text))
                    if rank > best_rank:
                        best_rank = rank
                        best = candidates[-1]
        findings["bands"].append(
            {
                "index": index,
                "rect": {"x0": band["x0"], "y0": band["y0"], "x1": band["x1"], "y1": band["y1"]},
                "color": band["color"],
                "pixels": band["pixels"],
                "best": best,
                "candidates": candidates,
            }
        )
    if not bands:
        for scale in (2, 3, 4):
            variant = preprocess_variants(img, scale)["binarized"]
            variant_path = os.path.join(outdir, f"{label}-full-{scale}x.png")
            variant.save(variant_path)
            ocr = ocr_words(variant_path)
            for line in ocr["lines"]:
                text = line["text"].strip()
                if not text:
                    continue
                findings["full_pane"].append(
                    {
                        "text": text,
                        "rect": {key: line[key] for key in ("x", "y", "w", "h")},
                        "matches": snap_to_source(text, source_entries),
                        "image": os.path.relpath(variant_path, ROOT),
                    }
                )
    return findings


def write_report(outdir, findings, dialog_lines, notes):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    payload = {
        "timestamp": timestamp,
        "dialog": dialog_lines,
        "notes": notes,
        "findings": findings,
    }
    with open(os.path.join(outdir, "report.json"), "w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2)

    lines = [f"# VBA error OCR report — {timestamp}", ""]
    if dialog_lines:
        lines.append("## Dialog")
        for item in dialog_lines:
            lines.append(f"- {item}")
        lines.append("")
    for note in notes:
        lines.append(f"- {note}")
    if notes:
        lines.append("")
    for finding in findings:
        lines.append(f"## {finding['label']} ({finding['image']})")
        if finding.get("module"):
            lines.append(f"Module: `{finding['module']}`")
        for band in finding.get("bands", []):
            rect = band["rect"]
            lines.append(
                f"### Highlighted band {band['index']} "
                f"(x={rect['x0']}, y={rect['y0']}, {rect['x1'] - rect['x0']}x{rect['y1'] - rect['y0']}, color={band['color']})"
            )
            best = band.get("best")
            if best:
                lines.append(
                    f"- OCR ({best['crop']} crop, {best['scale']}x {best['variant']}): `{best['text']}`"
                )
                for match in best["matches"]:
                    lines.append(
                        f"- Match {match['score']:.3f}: `{match['file']}:{match['line']}` — `{match['source']}`"
                    )
            else:
                lines.append("- OCR produced no text for this band; see candidate images.")
            lines.append("")
        for item in finding.get("full_pane", []):
            lines.append(f"- OCR: `{item['text']}`")
            for match in item["matches"][:1]:
                lines.append(
                    f"  - Match {match['score']:.3f}: `{match['file']}:{match['line']}` — `{match['source']}`"
                )
        lines.append("")
    report_path = os.path.join(outdir, "report.md")
    with open(report_path, "w", encoding="utf-8") as handle:
        handle.write("\n".join(lines))

    with open(REPORT_LOG, "a", encoding="utf-8") as handle:
        handle.write(f"{timestamp} | capture {os.path.basename(outdir)}\n")
        for item in dialog_lines:
            handle.write(f"{timestamp} | dialog | {item}\n")
        for finding in findings:
            for band in finding.get("bands", []):
                best = band.get("best")
                if not best:
                    continue
                top = best["matches"][0] if best["matches"] else None
                handle.write(f"{timestamp} | band {band['index']} | {best['text']}\n")
                if top:
                    handle.write(
                        f"{timestamp} | match {top['score']:.3f} | {top['file']}:{top['line']} | {top['source']}\n"
                    )
    return report_path


def capture_excel_context():
    """Best-effort COM attach to the running Excel to record where the error
    happened (workbook, sheet, selection). Wrapped in try/except because a
    modal dialog can block COM or another Excel instance may be active."""
    ctx = {}
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        try:
            wb = excel.ActiveWorkbook
            ctx["workbook"] = wb.Name if wb else None
        except Exception:
            ctx["workbook"] = None
        try:
            ws = excel.ActiveSheet
            ctx["sheet"] = ws.Name if ws else None
        except Exception:
            ctx["sheet"] = None
        try:
            sel = excel.Selection
            ctx["selection"] = sel.Address if sel else None
        except Exception:
            ctx["selection"] = None
        try:
            ctx["visible"] = bool(excel.Visible)
        except Exception:
            ctx["visible"] = None
    except Exception:
        ctx["error"] = "COM attach failed (modal dialog may block Excel)"
    return ctx


def capture_live(pid, outdir):
    """Attach to a running Excel and capture every visible VBA window.

    Returns a dict of captured artifacts keyed by kind. Does not kill Excel.
    """
    os.makedirs(outdir, exist_ok=True)
    windows = find_excel_windows(pid)
    captured = {
        "windows": windows["all"],
        "dialog": None,
        "vbe": [],
        "code": None,
        "desktop": None,
        "context": capture_excel_context(),
    }

    if windows["dialogs"]:
        d = windows["dialogs"][0]
        captured["dialog"] = {
            "hwnd": d[0],
            "title": d[1],
            "lines": d[2],
            "image": capture_window(d[0], os.path.join(outdir, "dialog.png")),
        }

    for hwnd, text in windows["vbe"]:
        entry = {"hwnd": hwnd, "title": text, "image": capture_window(hwnd, os.path.join(outdir, "vbe.png"))}
        captured["vbe"].append(entry)
        code_windows = find_code_windows(hwnd)
        if code_windows:
            chwnd, ccaption = code_windows[0]
            captured["code"] = {
                "hwnd": chwnd,
                "caption": ccaption,
                "module": module_from_caption(ccaption),
                "image": capture_window(chwnd, os.path.join(outdir, "vbe-code.png")),
            }

    # Full desktop context around Excel (other monitors, overlapping apps, etc.)
    captured["desktop"] = capture_desktop(os.path.join(outdir, "desktop.png"))

    return captured


def cmd_collect(args):
    enable_dpi_awareness()
    os.makedirs(OCR_ROOT, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    outdir = os.path.join(OCR_ROOT, stamp)
    os.makedirs(outdir, exist_ok=True)
    os.makedirs(TEMP_DIR, exist_ok=True)
    kill_excel()

    workbook = args.workbook or DEFAULT_WORKBOOK
    probe = os.path.join(TEMP_DIR, "ocr_probe.xlsm")
    if os.path.exists(probe):
        os.remove(probe)
    shutil.copy2(workbook, probe)

    shared = {
        "excel": None,
        "wb": None,
        "pid": None,
        "run_error": None,
        "run_done": False,
        "setup_error": None,
    }
    deadline = time.time() + max(5, args.timeout)

    def run_worker():
        pythoncom.CoInitialize()
        excel = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            shared["excel"] = excel
            excel.Visible = True
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Open(probe)
            shared["wb"] = wb
            shared["pid"] = win32process.GetWindowThreadProcessId(excel.Hwnd)[1]
            try:
                excel.VBE.MainWindow.Visible = True
                module = args.module
                if module:
                    wb.VBProject.VBComponents(module).CodeModule.CodePane.Show()
            except Exception as exc:
                shared["setup_error"] = str(exc)
            try:
                excel.Run(args.macro)
                shared["run_done"] = True
            except Exception as exc:
                shared["run_error"] = str(exc)
                shared["run_done"] = True
        finally:
            shared["run_done"] = True
            try:
                if excel:
                    excel.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()

    worker = threading.Thread(target=run_worker, daemon=True)
    worker.start()

    while time.time() < deadline and shared["pid"] is None and not shared["run_done"]:
        time.sleep(0.05)
    pid = shared["pid"]

    captured = {}
    if pid:
        while time.time() < deadline and not shared["run_done"]:
            windows = find_excel_windows(pid)
            if windows["dialogs"]:
                live = capture_live(pid, outdir)
                captured["dialog"] = live["dialog"]
                captured["vbe"] = live["vbe"]
                captured["code"] = live["code"]
                captured["desktop"] = live["desktop"]
                captured["context"] = live["context"]
                captured["dialog_title"] = live["dialog"]["title"] if live["dialog"] else None
                captured["dialog_lines"] = live["dialog"]["lines"] if live["dialog"] else []
                captured["module"] = live["code"]["module"] if live["code"] else None
                kill_excel()  # VBA's compile dialog cannot be dismissed from outside; kill Excel to unblock
                break
            time.sleep(0.05)

    kill_excel()

    findings = []
    module = captured.get("module") or args.module
    sources = args.source or discover_sources(module, ROOT)
    if captured.get("code") and captured["code"].get("image"):
        findings.append(
            analyze_image(captured["code"]["image"], module, sources, outdir, "code")
        )
    if not findings and captured.get("vbe"):
        vbe_img = captured["vbe"][0]["image"] if captured["vbe"] else None
        if vbe_img:
            findings.append(analyze_image(vbe_img, module, sources, outdir, "vbe"))

    dialog_lines = captured.get("dialog_lines") or []
    notes = []
    if captured.get("dialog_title"):
        notes.append(f"Dialog title: {captured['dialog_title']}")
    if captured.get("code") and captured["code"].get("caption"):
        notes.append(f"Code window: {captured['code']['caption']}")
    if captured.get("desktop"):
        notes.append(f"Desktop context captured: {os.path.basename(captured['desktop'])}")
    ctx = captured.get("context") or {}
    ctx_parts = []
    for key in ("workbook", "sheet", "selection", "visible"):
        if key in ctx and ctx[key] is not None:
            ctx_parts.append(f"{key}={ctx[key]}")
    if ctx_parts:
        notes.append("Excel context: " + ", ".join(ctx_parts))
    elif ctx.get("error"):
        notes.append(f"Excel context: {ctx['error']}")
    else:
        notes.append("Excel context: attempted (COM blocked by modal dialog or no active workbook)")
    if shared.get("run_error"):
        notes.append(f"Macro raised: {shared['run_error'][:400]}")
    if shared.get("setup_error"):
        notes.append(f"VBE setup issue: {shared['setup_error']}")
    if not captured.get("dialog_lines") and not captured.get("code"):
        notes.append("No error dialog or VBE code window was captured before the timeout.")

    report_path = write_report(outdir, findings, dialog_lines, notes)
    print(f"dialog: {' | '.join(dialog_lines) if dialog_lines else '(none)'}")
    for finding in findings:
        for band in finding.get("bands", []):
            best = band.get("best")
            if not best:
                continue
            print(f"band {band['index']} OCR: {best['text']}")
            if best["matches"]:
                top = best["matches"][0]
                print(f"  match {top['score']:.3f}: {top['file']}:{top['line']} — {top['source']}")
    print(f"report: {os.path.relpath(report_path, ROOT)}")
    if not args.keep_probe and os.path.exists(probe):
        try:
            os.remove(probe)
        except OSError:
            pass
    return 0


def cmd_ocr_live(args):
    """Attach to a running Excel and capture its live VBA UI + OCR."""
    enable_dpi_awareness()
    os.makedirs(OCR_ROOT, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    outdir = args.out or os.path.join(OCR_ROOT, f"live_{stamp}")
    os.makedirs(outdir, exist_ok=True)

    pid = args.pid
    if not pid:
        # Find the first Excel that has a visible dialog.
        for proc in psutil.process_iter(["name", "pid"]):
            if (proc.info["name"] or "").lower() != "excel.exe":
                continue
            windows = find_excel_windows(proc.info["pid"])
            if windows["dialogs"]:
                pid = proc.info["pid"]
                break
    if not pid:
        print("no running Excel with a visible dialog found")
        return 2

    pythoncom.CoInitialize()
    try:
        captured = capture_live(pid, outdir)
    finally:
        pythoncom.CoUninitialize()

    module = args.module
    if captured.get("code") and captured["code"].get("module"):
        module = captured["code"]["module"]
    sources = args.source or discover_sources(module, ROOT)

    findings = []
    if captured.get("code") and captured["code"].get("image"):
        findings.append(analyze_image(captured["code"]["image"], module, sources, outdir, "code"))
    if not findings and captured.get("vbe"):
        vbe_img = captured["vbe"][0]["image"] if captured["vbe"] else None
        if vbe_img:
            findings.append(analyze_image(vbe_img, module, sources, outdir, "vbe"))

    dialog = captured.get("dialog") or {}
    dialog_lines = dialog.get("lines") or []
    notes = []
    if dialog.get("title"):
        notes.append(f"Dialog title: {dialog['title']}")
    if captured.get("code") and captured["code"].get("caption"):
        notes.append(f"Code window: {captured['code']['caption']}")
    if captured.get("desktop"):
        notes.append(f"Desktop context captured: {os.path.basename(captured['desktop'])}")
    captured_windows = captured.get("windows") or []
    notes.append(f"Captured {len(captured_windows)} visible Excel/VBA windows")
    ctx = captured.get("context") or {}
    ctx_parts = []
    for key in ("workbook", "sheet", "selection", "visible"):
        if key in ctx and ctx[key] is not None:
            ctx_parts.append(f"{key}={ctx[key]}")
    if ctx_parts:
        notes.append("Excel context: " + ", ".join(ctx_parts))
    elif ctx.get("error"):
        notes.append(f"Excel context: {ctx['error']}")
    else:
        notes.append("Excel context: attempted (COM blocked by modal dialog or no active workbook)")

    report_path = write_report(outdir, findings, dialog_lines, notes)
    print(f"dialog: {' | '.join(dialog_lines) if dialog_lines else '(none)'}")
    for finding in findings:
        for band in finding.get("bands", []):
            best = band.get("best")
            if not best:
                continue
            print(f"band {band['index']} OCR: {best['text']}")
            if best["matches"]:
                top = best["matches"][0]
                print(f"  match {top['score']:.3f}: {top['file']}:{top['line']} — {top['source']}")
    print(f"report: {os.path.relpath(report_path, ROOT)}")
    return 0


def cmd_ocr_all(args):
    """Batch-OCR every screenshot already under log\\ocr\\ into one combined report.

    Default is a fast summary of existing report.json files (no re-OCR). Pass
    --refresh to re-run the (improved) OCR pipeline on every captured code pane.
    """
    enable_dpi_awareness()
    if not os.path.isdir(OCR_ROOT):
        print(f"no {OCR_ROOT} directory yet")
        return 2
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    outdir = args.out or os.path.join(OCR_ROOT, f"all_{stamp}")
    os.makedirs(outdir, exist_ok=True)

    findings = []
    dialog_lines = []
    notes = []
    dirs = sorted(
        [d for d in os.listdir(OCR_ROOT) if os.path.isdir(os.path.join(OCR_ROOT, d))],
        reverse=True,
    )
    for name in dirs[:50]:
        d = os.path.join(OCR_ROOT, name)
        report_json = os.path.join(d, "report.json")
        report_md = os.path.join(d, "report.md")
        code_img = None
        for cand in ("vbe-code.png", "dialog.png"):
            p = os.path.join(d, cand)
            if os.path.isfile(p):
                code_img = p
                break

        if args.refresh and code_img and code_img.endswith("vbe-code.png"):
            module = args.module
            sources = args.source or discover_sources(module, ROOT)
            f = analyze_image(code_img, module, sources, outdir, f"batch-{name}")
            f["label"] = f"batch {name}"
            findings.append(f)
        elif os.path.isfile(report_json):
            try:
                with open(report_json, "r", encoding="utf-8", errors="replace") as h:
                    payload = json.load(h)
                findings.append(
                    {
                        "label": f"batch {name}",
                        "image": os.path.relpath(d, ROOT),
                        "module": payload.get("exact", {}).get("module")
                        or (payload.get("findings") or [{}])[0].get("module"),
                        "bands": [],
                        "full_pane": [],
                        "_summary": payload,
                    }
                )
                for item in payload.get("dialog", []):
                    dialog_lines.append(f"{name}: {item}")
                for note in payload.get("notes", []):
                    notes.append(f"{name}: {note}")
            except (OSError, ValueError):
                pass
        if os.path.isfile(report_md):
            try:
                with open(report_md, "r", encoding="utf-8", errors="replace") as h:
                    head = h.read(500)
                for line in head.splitlines():
                    if "Variable not defined" in line or "Compile error" in line:
                        dialog_lines.append(f"{name}: {line.strip()}")
            except OSError:
                pass

    if args.refresh:
        notes.append(f"Re-OCR'd {len(findings)} code panes across {len(dirs)} capture directories")
    else:
        notes.append(f"Fast-summarized {len(dirs)} capture directories (no re-OCR); use --refresh to re-OCR")
    report_path = write_report(outdir, findings, dialog_lines, notes)
    print(f"batch directories: {len(dirs)}")
    print(f"report: {os.path.relpath(report_path, ROOT)}")
    return 0


def cmd_analyze(args):
    enable_dpi_awareness()
    image_path = args.image
    if not os.path.isabs(image_path):
        image_path = os.path.join(ROOT, image_path)
    if not os.path.isfile(image_path):
        print(f"missing image: {image_path}")
        return 2
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    outdir = args.out or os.path.join(OCR_ROOT, stamp)
    os.makedirs(outdir, exist_ok=True)
    module = args.module
    sources = args.source or discover_sources(module, ROOT)
    findings = [analyze_image(image_path, module, sources, outdir, "analyze")]
    report_path = write_report(outdir, findings, args.dialog or [], [])
    print(f"report: {os.path.relpath(report_path, ROOT)}")
    return 0


def main():
    parser = argparse.ArgumentParser(description="Capture and OCR VBA UI errors.")
    subparsers = parser.add_subparsers(dest="command")

    collect = subparsers.add_parser("collect", help="Run a macro, capture the error UI, then OCR it.")
    collect.add_argument("--macro", required=True)
    collect.add_argument("--module")
    collect.add_argument("--workbook")
    collect.add_argument("--source", nargs="*", help="Module source files used for OCR snapping.")
    collect.add_argument("--timeout", type=int, default=25)
    collect.add_argument("--keep-probe", action="store_true")

    live = subparsers.add_parser("ocr-live", help="Attach to running Excel and capture live VBA UI.")
    live.add_argument("--pid", type=int, help="Excel process id (auto-detected if omitted).")
    live.add_argument("--module")
    live.add_argument("--source", nargs="*")
    live.add_argument("--out")

    allcmd = subparsers.add_parser("ocr-all", help="Batch-OCR all screenshots under log\\ocr\\.")
    allcmd.add_argument("--module")
    allcmd.add_argument("--source", nargs="*")
    allcmd.add_argument(
        "--refresh",
        action="store_true",
        help="Re-run the OCR pipeline on every captured code pane (slower).",
    )
    allcmd.add_argument("--out")

    analyze = subparsers.add_parser("analyze", help="OCR an existing screenshot.")
    analyze.add_argument("image")
    analyze.add_argument("--module")
    analyze.add_argument("--source", nargs="*")
    analyze.add_argument("--dialog", nargs="*")
    analyze.add_argument("--out")

    args = parser.parse_args()
    if args.command == "collect":
        return cmd_collect(args)
    if args.command == "ocr-live":
        return cmd_ocr_live(args)
    if args.command == "ocr-all":
        return cmd_ocr_all(args)
    if args.command == "analyze":
        return cmd_analyze(args)
    parser.print_help()
    return 1


if __name__ == "__main__":
    sys.exit(main())
