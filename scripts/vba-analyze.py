#!/usr/bin/env python3
"""Static VBA analysis for the Pending Pickup template.

Scans Pending Pickup Template.xlsm with oletools VBA_Parser and reports:
  - whether VBA macros are present
  - VBA stomping (source/p-code mismatch - suspicious)
  - XLM (Excel 4.0) macros - clearly malicious if present
  - per-macro suspicious keyword hits (Shell, CreateObject, ExecuteExcel4LM,
    AutoExec, AutoOpen, SendKeys, Environ, WScript.Shell)

Usage:
  python scripts/vba-analyze.py
  python scripts/vba-analyze.py --template "Pending Pickup Template.xlsm"

Exits non-zero only on clear problems (XLM macros or a parse error). Stomping
and keyword hits are reported as warnings so they are visible but do not break
the pipeline on legitimate business-macro patterns.
"""
from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path

from oletools.olevba import VBA_Parser

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_TEMPLATE = ROOT / "Pending Pickup Template.xlsm"
LOG_DIR = ROOT / "log"
LOG_DIR.mkdir(parents=True, exist_ok=True)
ANALYZE_LOG = LOG_DIR / "vba-analyze.log"

SUSPICIOUS_KEYWORDS = (
    "AutoExec",
    "AutoOpen",
    "Workbook_Open",
    "Worksheet_Activate",
    "Shell",
    "CreateObject",
    "GetObject",
    "ExecuteExcel4LM",
    "SendKeys",
    "Environ",
    "WScript.Shell",
    "MSXML2.XMLHTTP",
    "WinHttp.WinHttpRequest",
    "cmd.exe",
    "powershell",
)


def log(msg: str) -> None:
    line = f"{__import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {msg}"
    print(line)
    with ANALYZE_LOG.open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def count_keyword_hits(vba_code: str) -> dict[str, int]:
    hits = {}
    for kw in SUSPICIOUS_KEYWORDS:
        n = len(re.findall(re.escape(kw), vba_code, re.IGNORECASE))
        if n:
            hits[kw] = n
    return hits


def analyze(template: Path) -> int:
    if not template.exists():
        log(f"ERROR: template not found: {template}")
        return 2

    log(f"=== VBA static analysis: {template.name} ===")
    try:
        vp = VBA_Parser(str(template))
    except Exception as e:
        log(f"ERROR: could not open template for VBA analysis: {e}")
        return 2

    try:
        has_macros = vp.detect_vba_macros()
        stomping = vp.detect_vba_stomping()
        xlm = vp.detect_xlm_macros()

        log(f"VBA macros present: {has_macros}")
        log(f"VBA stomping (source/p-code mismatch): {stomping}")
        log(f"XLM (Excel 4.0) macros: {xlm}")

        macros = list(vp.extract_all_macros())
        log(f"Extracted macros: {len(macros)}")

        keyword_total = 0
        for filename, stream_path, vba_filename, vba_code in macros:
            if not vba_code:
                continue
            hits = count_keyword_hits(vba_code)
            if hits:
                keyword_total += sum(hits.values())
                detail = ", ".join(f"{k}={v}" for k, v in sorted(hits.items()))
                log(f"  suspicious keywords in {vba_filename}: {detail}")

        if xlm:
            log("RESULT: FAIL - XLM (Excel 4.0) macros detected")
            return 1

        if stomping:
            log("WARNING: VBA stomping detected - VBA source and compiled p-code differ")

        if keyword_total:
            log(f"WARNING: {keyword_total} suspicious keyword hit(s) across macros")

        log("RESULT: OK - no clear VBA threats detected")
        return 0
    finally:
        vp.close()


def main() -> int:
    ap = argparse.ArgumentParser(description="Statically analyze VBA in the Pending Pickup template")
    ap.add_argument("--template", default=str(DEFAULT_TEMPLATE), help="Path to the .xlsm template")
    args = ap.parse_args()
    return analyze(Path(args.template))


if __name__ == "__main__":
    sys.exit(main())
