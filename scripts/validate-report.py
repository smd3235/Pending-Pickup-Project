#!/usr/bin/env python3
"""Validate the latest (or specified) Pending Pickup report workbook in depth.

Checks beyond basic log scanning:
- workbook open/readability, expected sheets (Dashboard, Raw Data, Audit Log, manager tabs)
- Raw Data headers and row count, no fully blank header cells in used range
- Dashboard title/KPI presence
- Audit Log key fields and numeric consistency where parseable
- per-manager tab structure (header row present, has data or consistent empty state)
- screenshot coverage for report sheets (report-*.png) and template screenshots
- output vs completion marker and source expectations where inferable

Usage:
  python scripts/validate-report.py
  python scripts/validate-report.py --report "Reports/....xlsx"
  python scripts/validate-report.py --json
  python scripts/validate-report.py --strict
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path

try:
    import openpyxl
except Exception:
    openpyxl = None

ROOT = Path(__file__).resolve().parent.parent
REPORTS_DIR = ROOT / "Reports"
LOG_DIR = ROOT / "log"
DOC_SHOTS = ROOT / "documentation" / "screenshots"

RAW_EXPECTED_HEADERS = {"Name", "EMPI", "Treatment Status", "Assigned Sales Manager", "AUDIT: TP End Date"}
MANAGER_HEADER_ROW = 16  # matches VBA constant


def add(findings, severity, source, message, extra=None):
    findings.append({"severity": severity, "source": str(source), "message": message, "extra": extra or {}})


def find_latest_report(explicit=None):
    if explicit:
        p = Path(explicit)
        if not p.is_absolute():
            p = ROOT / p
        return p if p.exists() else None
    if not REPORTS_DIR.exists():
        return None
    cands = [p for p in REPORTS_DIR.iterdir() if p.is_file() and p.suffix.lower() == ".xlsx" and not p.name.startswith("~$")]
    return max(cands, key=lambda p: p.stat().st_mtime) if cands else None


def slugify_sheet(name: str) -> str:
    s = name.lower()
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    return s or "sheet"


def parse_audit_log(ws):
    data = {}
    # read label/value pairs in first ~40 rows, cols A-B
    for r in range(1, 45):
        k = ws.cell(r, 1).value
        v = ws.cell(r, 2).value
        if k is None:
            continue
        ks = str(k).strip().rstrip(":")
        data[ks] = v
    return data


def to_int(v):
    try:
        if v is None or v == "":
            return None
        return int(float(v))
    except Exception:
        return None


def main():
    ap = argparse.ArgumentParser(description="Deep-validate latest Pending Pickup report workbook")
    ap.add_argument("--report", default=None)
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--strict", action="store_true")
    args = ap.parse_args()

    findings = []
    report = find_latest_report(args.report)
    if report is None:
        add(findings, "error", REPORTS_DIR, "No report .xlsx found")
    elif not report.exists():
        add(findings, "error", report, "Report path does not exist")

    if openpyxl is None:
        add(findings, "error", "openpyxl", "openpyxl is required for report validation")
        result = {"ok": False, "report": str(report) if report else None, "counts": {"error": 1, "warn": 0, "info": 0}, "findings": findings}
        print(json.dumps(result, indent=2) if args.json else "openpyxl missing")
        return 1

    if report and report.exists():
        try:
            size = report.stat().st_size
            if size < 10_000:
                add(findings, "warn", report, f"Report file is very small ({size} bytes)")
        except Exception as e:
            add(findings, "warn", report, f"Could not stat report: {e}")

        try:
            wb = openpyxl.load_workbook(report, read_only=False, data_only=True)
        except Exception as e:
            add(findings, "error", report, f"Could not open workbook: {e}")
            wb = None

        if wb is not None:
            try:
                names = wb.sheetnames
                nset = set(names)
                for req in ("Dashboard", "Raw Data", "Audit Log"):
                    if req not in nset:
                        add(findings, "error", report, f"Missing required sheet: {req}", {"sheets": names})
                manager_tabs = [n for n in names if n not in {"Dashboard", "Raw Data", "Audit Log"}]
                if not manager_tabs:
                    add(findings, "warn", report, "No manager tabs found", {"sheets": names})

                # Dashboard checks
                if "Dashboard" in nset:
                    ws = wb["Dashboard"]
                    texts = []
                    for r in range(1, 8):
                        for c in range(1, 20):
                            v = ws.cell(r, c).value
                            if v is not None and str(v).strip():
                                texts.append(str(v).strip())
                    blob = " ".join(texts).upper()
                    if "PENDING PICKUP" not in blob:
                        add(findings, "warn", report, "Dashboard title/header text not found in top rows")
                    # expect at least some KPI-like values
                    if not any(re.search(r"\d", t) or t == "N/A" for t in texts):
                        add(findings, "warn", report, "Dashboard appears to have no numeric/KPI values near top")

                # Raw Data checks
                raw_rows = 0
                raw_headers = []
                if "Raw Data" in nset:
                    ws = wb["Raw Data"]
                    raw_rows = ws.max_row or 0
                    raw_cols = ws.max_column or 0
                    for c in range(1, min(raw_cols, 80) + 1):
                        v = ws.cell(1, c).value
                        raw_headers.append("" if v is None else str(v).strip())
                    header_set = {h for h in raw_headers if h}
                    miss = sorted(RAW_EXPECTED_HEADERS - header_set)
                    if miss:
                        add(findings, "warn", report, f"Raw Data missing expected headers: {', '.join(miss)}", {"headers": [h for h in raw_headers if h][:40]})
                    if raw_rows < 2:
                        add(findings, "warn", report, "Raw Data has no data rows")
                    # basic blank-header check inside used range
                    used_hdr = [h for h in raw_headers if h]
                    if used_hdr and len(used_hdr) < max((i for i, h in enumerate(raw_headers, 1) if h), default=0):
                        add(findings, "info", report, "Raw Data has gaps in header row (may be expected)")

                # Audit Log checks and cross-checks
                audit = {}
                if "Audit Log" in nset:
                    audit = parse_audit_log(wb["Audit Log"])
                    blob = " ".join(str(k) for k in audit.keys())
                    for key in ("Report Generated", "Source File", "Total Records in Source", "Pending Pickup Records"):
                        if key not in audit and key.lower() not in blob.lower():
                            add(findings, "warn", report, f"Audit Log missing field: {key}")
                    pending = to_int(audit.get("Pending Pickup Records"))
                    total = to_int(audit.get("Total Records in Source"))
                    if pending is not None and raw_rows and raw_rows > 0:
                        # raw data usually header + pending rows
                        expected = pending + 1
                        if raw_rows != expected:
                            add(findings, "warn", report, f"Raw Data row count ({raw_rows}) != Audit pending+1 ({expected})", {"audit_pending": pending, "raw_rows": raw_rows})
                    if pending is not None and total is not None and pending > total:
                        add(findings, "error", report, "Audit: Pending Pickup Records exceeds Total Records in Source", audit)

                # Manager tab checks
                for m in manager_tabs:
                    ws = wb[m]
                    # header around MANAGER_HEADER_ROW
                    hdr_vals = [ws.cell(MANAGER_HEADER_ROW, c).value for c in range(1, 20)]
                    hdr_nonempty = [str(v).strip() for v in hdr_vals if v is not None and str(v).strip()]
                    if len(hdr_nonempty) < 3:
                        add(findings, "warn", report, f"Manager tab '{m}' missing/weak headers at row {MANAGER_HEADER_ROW}", {"headers": hdr_nonempty})
                    # ensure sheet not completely empty
                    if (ws.max_row or 0) < MANAGER_HEADER_ROW:
                        add(findings, "warn", report, f"Manager tab '{m}' has very few rows ({ws.max_row})")

                # Screenshot coverage
                if DOC_SHOTS.exists():
                    # report screenshots expected for dashboard/raw data/audit log and each manager tab
                    expected = {"dashboard": DOC_SHOTS / "report-dashboard.png", "raw data": DOC_SHOTS / "report-raw-data.png", "audit log": DOC_SHOTS / "report-audit-log.png"}
                    for m in manager_tabs:
                        expected[m] = DOC_SHOTS / f"report-{slugify_sheet(m)}.png"
                    missing = [name for name, p in expected.items() if not p.exists()]
                    if missing:
                        add(findings, "warn", report, f"Missing report screenshot(s): {', '.join(sorted(missing)[:8])}" + ("..." if len(missing) > 8 else ""), {"count": len(missing)})
                    tiny = [name for name, p in expected.items() if p.exists() and p.stat().st_size < 1024]
                    if tiny:
                        add(findings, "warn", report, f"Very small report screenshot(s): {', '.join(sorted(tiny)[:8])}")

                # Completion marker cross-check
                marker = LOG_DIR / "pending-pickup-complete.txt"
                if marker.exists():
                    try:
                        mtxt = marker.read_text(encoding="utf-8", errors="replace")
                        mm = re.search(r"saved=(.+)", mtxt)
                        if mm:
                            saved = Path(mm.group(1).strip())
                            if saved.exists() and saved.resolve() != report.resolve():
                                add(findings, "info", marker, "Completion marker points to different report than validated report", {"marker_saved": str(saved), "validated": str(report)})
                    except Exception:
                        pass
            finally:
                try:
                    wb.close()
                except Exception:
                    pass

    counts = {"error": 0, "warn": 0, "info": 0}
    for f in findings:
        counts[f["severity"]] = counts.get(f["severity"], 0) + 1
    ok = counts.get("error", 0) == 0 and (not args.strict or counts.get("warn", 0) == 0)
    result = {"ok": ok, "report": str(report) if report else None, "counts": counts, "findings": findings}

    if args.json:
        print(json.dumps(result, indent=2))
    else:
        print(f"Report: {report if report else '(none)'}")
        print(f"Validation: {counts.get('error',0)} error(s), {counts.get('warn',0)} warning(s), {counts.get('info',0)} info")
        for f in findings:
            print(f"[{f['severity'].upper()}] {f['source']} - {f['message']}")
            ex = f.get("extra") or {}
            if ex:
                # keep compact
                s = json.dumps(ex, ensure_ascii=False)
                if len(s) > 300:
                    s = s[:300] + "..."
                print(f"  extra: {s}")
        print("OK: report validation passed." if ok else "ISSUES DETECTED in report validation.")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
