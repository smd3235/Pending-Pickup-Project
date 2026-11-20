#!/usr/bin/env python3
"""Shared, dependency-light helpers for VBA error OCR collection.

Pure text/image processing only: no win32com/pythoncom/psutil at import time,
so modules that import this stay fast (important for the sub-30s pytest gate).
"""
from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import difflib
import os
import re
from pathlib import Path

user32 = ctypes.windll.user32


def normalize(text):
    text = re.sub(r"^\s*\d+\s+", "", text or "")
    text = text.lower()
    text = re.sub(r"[^a-z0-9_ ]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def load_sources(paths):
    entries = []
    for path in paths:
        if not path or not os.path.isfile(path):
            continue
        try:
            with open(path, "r", encoding="utf-8", errors="replace") as handle:
                for number, line in enumerate(handle, start=1):
                    entries.append((path, number, line.rstrip("\r\n")))
        except OSError:
            continue
    return entries


def discover_sources(module, root):
    vba_dir = root / "vba"
    if module:
        candidate = vba_dir / f"{module}.bas"
        return [candidate] if candidate.is_file() else []
    if vba_dir.is_dir():
        return [p for p in sorted(vba_dir.iterdir()) if p.suffix.lower() == ".bas"]
    return []


def snap_to_source(text, source_entries, top=3):
    needle = normalize(text)
    if len(needle) < 3 or not source_entries:
        return []
    needle_compact = needle.replace(" ", "")
    scored = []
    for path, number, source in source_entries:
        haystack = normalize(source)
        if len(haystack) < 3:
            continue
        ratio = difflib.SequenceMatcher(None, needle, haystack).ratio()
        compact_ratio = difflib.SequenceMatcher(None, needle_compact, haystack.replace(" ", "")).ratio()
        ratio = max(ratio, compact_ratio)
        if needle in haystack and len(needle) >= 6:
            ratio = max(ratio, 0.99)
        scored.append((ratio, path, number, source))
    scored.sort(key=lambda item: item[0], reverse=True)
    matches = []
    for ratio, path, number, source in scored[:top]:
        matches.append(
            {
                "score": round(ratio, 3),
                "file": os.path.relpath(path, Path(os.getcwd())),
                "line": number,
                "source": source.strip(),
            }
        )
    return matches


def detect_bands(img, min_run=30, max_bands=6):
    width, height = img.size
    pixels = img.load()
    rows = []
    for y in range(height):
        counts = {}
        for x in range(0, width, 2):
            r, g, b = pixels[x, y][:3]
            if r > 235 and g > 235 and b > 235:
                continue
            if abs(r - g) < 18 and abs(g - b) < 18 and (r > 90 or r < 40):
                continue
            counts[(r, g, b)] = counts.get((r, g, b), 0) + 1
        if not counts:
            continue
        color, count = max(counts.items(), key=lambda item: item[1])
        if count * 2 >= min_run and (max(color) - min(color) >= 30 or sum(color) < 300):
            rows.append((y, color, count * 2))
    bands = []
    for y, color, count in rows:
        if bands and y - bands[-1]["y1"] <= 3:
            bands[-1]["y1"] = y
            bands[-1]["pixels"] += count
            bands[-1]["colors"][color] = bands[-1]["colors"].get(color, 0) + count
        else:
            bands.append({"y0": y, "y1": y, "pixels": count, "colors": {color: count}})
    result = []
    for band in bands:
        band_height = band["y1"] - band["y0"] + 1
        if band_height < 5 or band_height > 220:
            continue
        color = max(band["colors"].items(), key=lambda item: item[1])[0]
        xs = []
        for y in range(band["y0"], band["y1"] + 1):
            for x in range(0, width, 2):
                if pixels[x, y][:3] == color:
                    xs.append(x)
        if not xs:
            continue
        result.append(
            {
                "y0": band["y0"],
                "y1": band["y1"],
                "x0": max(0, min(xs) - 16),
                "x1": min(width, max(xs) + 32),
                "height": band_height,
                "color": color,
                "pixels": band["pixels"],
            }
        )
    result.sort(key=lambda item: item["pixels"], reverse=True)
    return result[:max_bands]


def flatten_selection(img, color, tolerance=90):
    from PIL import Image

    width, height = img.size
    pixels = img.load()
    flat = Image.new("L", (width, height), 255)
    target = flat.load()
    for y in range(height):
        for x in range(width):
            r, g, b = pixels[x, y][:3]
            distance = abs(r - color[0]) + abs(g - color[1]) + abs(b - color[2])
            target[x, y] = 255 if distance < tolerance else 0
    return flat


def preprocess_variants(img, scale):
    from PIL import Image, ImageFilter, ImageOps

    width, height = img.size
    resized = img.resize((width * scale, height * scale), Image.LANCZOS)
    gray = ImageOps.autocontrast(resized.convert("L"))
    sharp = gray.filter(ImageFilter.SHARPEN)
    binary = sharp.point(lambda value: 255 if value > 150 else 0)
    smooth_binary = binary.filter(ImageFilter.MedianFilter(3))
    unsharp = gray.filter(ImageFilter.UnsharpMask(radius=2, percent=150, threshold=3))
    inverted = ImageOps.invert(gray)
    return {
        "gray": sharp,
        "binarized": binary,
        "smooth-binary": smooth_binary,
        "unsharp": unsharp,
        "inverted-gray": inverted,
    }


def dialog_child_texts(hwnd):
    texts = []

    def cb(child, _):
        try:
            cls = ctypes.windll.user32.GetClassNameW(child)
            buf = ctypes.create_unicode_buffer(1024)
            ctypes.windll.user32.GetWindowTextW(child, buf, 1024)
            text = buf.value.strip()
        except Exception:
            return True
        if text and cls in ("Static", "Button", "Edit", "SysLink"):
            if text not in texts:
                texts.append(text)
        return True

    try:
        ctypes.windll.user32.EnumChildWindows(
            hwnd, ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)(cb), 0
        )
    except Exception:
        pass
    return texts
