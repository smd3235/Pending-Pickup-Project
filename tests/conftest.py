"""Pytest configuration for Pending Pickup project tests.

Adds the scripts/ directory to sys.path so tests can import the pure-Python
helper modules (vba-error-ocr.py, test-compile.py) without launching Excel.
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))
