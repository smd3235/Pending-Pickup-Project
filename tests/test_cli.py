"""Smoke tests for CLI entry points (no Excel required)."""
from __future__ import annotations

import subprocess
import sys


def _run(args):
    return subprocess.run(
        [sys.executable, *args],
        capture_output=True,
        text=True,
        timeout=30,
    )


def test_vba_error_ocr_help_returns_zero():
    res = _run(["scripts/vba-error-ocr.py", "--help"])
    assert res.returncode == 0, res.stderr
    assert "collect" in res.stdout or "analyze" in res.stdout


def test_ocr_analyze_missing_file_nonzero():
    res = _run(["scripts/vba-error-ocr.py", "analyze", "scripts/temp/does-not-exist.png"])
    assert res.returncode != 0
