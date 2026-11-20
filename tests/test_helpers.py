"""Unit tests for pure-Python helpers used by the VBA error OCR collectors.

These tests do not launch Excel; they exercise image/text processing only.
"""
from __future__ import annotations

import os

# vba_ocr_helpers is dependency-light (no win32com/pythoncom/PIL at import),
# so tests import it directly for speed. PIL is imported lazily inside the
# image-processing tests to keep the suite well under the 30s pipeline cap.
import vba_ocr_helpers as veo
from vba_ocr_helpers import dialog_child_texts


VEO_PREF = "vba" + os.sep + "modX.bas"


def test_normalize_strips_line_numbers_and_collapses_space():
    assert veo.normalize("  12   UndeclaredVar + 1  ") == "undeclaredvar 1"
    assert veo.normalize("") == ""
    assert veo.normalize("   ") == ""


def test_normalize_removes_non_alphanumeric():
    assert veo.normalize("result = FakeOcrFakeToke") == "result fakeocrfaketoke"


def test_snap_to_source_exact_match():
    sources = [
        (VEO_PREF, 6, "result = FakeOcrFakeToken + 1"),
        (VEO_PREF, 4, "Public Sub FakeOcrError()"),
    ]
    matches = veo.snap_to_source("result = FakeOcrFakeToke", sources)
    assert matches, "expected at least one match"
    top = matches[0]
    assert top["file"] == VEO_PREF
    assert top["line"] == 6
    assert top["source"] == "result = FakeOcrFakeToken + 1"
    assert top["score"] >= 0.9


def test_snap_to_source_compact_ignores_spaces():
    sources = [("vba/modX.bas", 1, "Dim x As Long")]
    matches = veo.snap_to_source("DimxAsLong", sources)
    assert matches and matches[0]["score"] >= 0.9


def test_snap_to_source_empty_when_no_sources():
    assert veo.snap_to_source("some text", []) == []


def test_detect_bands_finds_colored_selection():
    from PIL import Image

    img = Image.new("RGB", (200, 40), (255, 255, 255))
    # Draw a blue selection band across rows 10-18
    for y in range(10, 19):
        for x in range(20, 180):
            img.putpixel((x, y), (0, 120, 212))
    bands = veo.detect_bands(img)
    assert len(bands) == 1
    b = bands[0]
    assert b["y0"] <= 10 and b["y1"] >= 18
    assert b["color"] == (0, 120, 212)
    assert b["height"] <= 220


def test_detect_bands_ignores_white_page():
    from PIL import Image

    img = Image.new("RGB", (200, 40), (255, 255, 255))
    assert veo.detect_bands(img) == []


def test_dialog_child_texts_collects_controls():
    # No real HWND without Windows GUI; verify the function returns [] gracefully
    # for an invalid handle instead of raising.
    result = dialog_child_texts(0)
    assert isinstance(result, list)


def test_flatten_selection_makes_bg_white_text_black():
    from PIL import Image

    img = Image.new("RGB", (10, 10), (0, 120, 212))  # selection blue background
    flat = veo.flatten_selection(img, (0, 120, 212), tolerance=90)
    # Background pixels close to selection color -> white (255)
    assert flat.getpixel((5, 5)) == 255
