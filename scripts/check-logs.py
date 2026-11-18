#!/usr/bin/env python3
"""Scan Pending Pickup run artifacts and logs for issues.

Checks:
- log/ files for FAILED/ERROR/Err=/exception/traceback/timeout/RPC/unavailable/warning patterns
- Excel lock/temp files (~$...) accidentally treated as real outputs in logs or Reports/
- latest/newest report workbook integrity and expected structure (openpyxl)
- completion marker consistency with newest report
- screenshot/report artifact freshness and presence
- run-report Python log health (outcome, completion marker, report found, no stall/kill)

Usage:
  python scripts/check-logs.py
  python scripts/check-logs.py --latest
  python scripts/check-logs.py --report "Reports/Pending Pickup ....xlsx"
  python scripts/check-logs.py --json
  python scripts/check-logs.py --strict   # exit 1 on warnings too
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

try:
    import openpyxl
except Exception:  # pragma: no cover
    openpyxl = None

ROOT = Path(__file__).resolve().parent.parent
LOG_DIR = ROOT / "log"
REPORTS_DIR = ROOT / "Reports"
DOC_SHOTS = ROOT / "documentation" / "screenshots"

# (pattern, severity, label)
LOG_PATTERNS = [
    (re.compile(r"\bFAILED\b|\bKILLED\b|\bERROR\b|\bErr=\s*[1-9]", re.I), "error", "failure/error"),
    (re.compile(r"Traceback \(most recent call last\)|pywintypes\.com_error|Exception occurred", re.I), "error", "exception/traceback"),
    (re.compile(r"The RPC server is unavailable|The remote procedure call failed", re.I), "error", "rpc/com failure"),
    (re.compile(r"\btimeout\b|timed out|exceeded .* without a new report file", re.I), "error", "timeout/stall"),
    (re.compile(r"VBA ERROR DIALOG|compile/runtime error", re.I), "error", "vba dialog/error"),
    (re.compile(r"Completion marker was not written", re.I), "error", "missing completion marker"),
    (re.compile(r"macro result\s*=\s*(?!True\b)\S+", re.I), "warn", "unexpected macro result"),
    (re.compile(r"\bWARNING\b", re.I), "warn", "warning"),
    (re.compile(r"not found|could not|unavailable|skipped", re.I), "warn", "missing/unavailable/skipped"),
]

REQUIRED_SHEETS = {"Dashboard", "Raw Data"}
OPTIONAL_SHEETS = {"Audit Log"}
RAW_EXPECTED_HEADERS_SUBSET = {"Name", "EMPI", "Treatment Status", "Assigned Sales Manager", "AUDIT: TP End Date"}


def add(findings, severity, source, message, line=None, extra=None):
    findings.append({
        "severity": severity,
        "source": str(source),
        "line": line,
        "message": message,
        "extra": extra or {},
    })


def iter_log_files(include_all=True):
    if not LOG_DIR.exists():
        return []
    files = [p for p in LOG_DIR.iterdir() if p.is_file()]
    if include_all:
        return sorted(files, key=lambda p: p.stat().st_mtime)
    keep = set()
    for name in ("macro-debug.log", "pending-pickup-complete.txt", "answer-dialogs.log", "compile-test.log", "watcher.log"):
        p = LOG_DIR / name
        if p.exists():
            keep.add(p)
    runs = sorted(LOG_DIR.glob("pending-pickup_*.log"), key=lambda p: p.stat().st_mtime)
    if runs:
        keep.add(runs[-1])
    return sorted(keep, key=lambda p: p.stat().st_mtime)


def scan_logs(findings, latest_only=False):
    files = iter_log_files(include_all=not latest_only)
    if not files:
        add(findings, "warn", LOG_DIR, "No log files found")
        return

    for path in files:
        try:
            text = path.read_text(encoding="utf-8", errors="replace")
        except Exception as e:
            add(findings, "error", path, f"Could not read log: {e}")
            continue
        lines = text.splitlines()

        for i, line in enumerate(lines, 1):
            if "~$Pending Pickup" in line and "REPORT OK" in line:
                add(findings, "warn", path, "Excel lock/temp file (~$...) included in REPORT OK log line", i, {"text": line.strip()[:300]})
            if "~$Pending Pickup" in line and "New report file" in line:
                add(findings, "warn", path, "Excel lock/temp file (~$...) included in new-report log line", i, {"text": line.strip()[:300]})

        report_ok = [ln for ln in lines if "REPORT OK:" in ln]
        if len(report_ok) > 3:
            add(findings, "warn", path, f"Repeated REPORT OK log lines ({len(report_ok)}); watchdog may be spamming", extra={"sample": report_ok[-1][:300]})

        for i, line in enumerate(lines, 1):
            if "Email/template sync check: email config matches no-macro" in line:
                continue
            if "Run call returned without COM error" in line:
                continue
            if "could not" in line.lower() and "VBAWarnings" in line:
                add(findings, "warn", path, "Could not set VBAWarnings registry", i, {"text": line.strip()[:300]})
                continue
            if "Could not set trusted location" in line:
                add(findings, "warn", path, "Could not set trusted location registry", i, {"text": line.strip()[:300]})
                continue
            for pat, sev, label in LOG_PATTERNS:
                if pat.search(line):
                    # benign compile-test cleanup: lingering Excel kill after probe run is expected
                    if "Killed lingering Excel pid" in line and "RESULT: COMPILE" not in line:
                        # keep visible as info-level note but not a failure
                        if "compile-test.log" in path.name or "test-compile" in line:
                            continue
                    if label == "missing/unavailable/skipped":
                        if "MISSING PICKUP DATE" in line and "label=" in line:
                            continue
                        if "Audit sheet written" in line:
                            continue
                        if "no-macro unreadable/missing" in line:
                            add(findings, "warn", path, "No-macro template unreadable/missing during sync check", i, {"text": line.strip()[:300]})
                            break
                        if "Completion marker was not written" in line:
                            add(findings, "error", path, "Completion marker was not written", i, {"text": line.strip()[:300]})
                            break
                        if path.name.startswith("pending-pickup_") and "Email/template sync skipped" in line:
                            add(findings, "warn", path, "Template/email sync skipped", i, {"text": line.strip()[:300]})
                            break
                        continue
                    if label == "failure/error" and "REPORT OK" in line:
                        continue
                    # benign cleanup lines
                    if "Killed lingering Excel pid" in line:
                        continue
                    if label == "warning" and path.name == "compile-test.log":
                        continue
                    add(findings, sev, path, f"Matched {label}", i, {"text": line.strip()[:350]})
                    break

    runs = sorted(LOG_DIR.glob("pending-pickup_*.log"), key=lambda p: p.stat().st_mtime)
    if runs:
        latest = runs[-1]
        txt = latest.read_text(encoding="utf-8", errors="replace")
        if "Final outcome: OK" not in txt and "run finished (outcome=OK)" not in txt:
            add(findings, "error", latest, "Latest run log does not show successful outcome")
        if "Completion marker contents:" not in txt and "Completion marker found" not in txt:
            add(findings, "warn", latest, "Latest run log does not clearly confirm completion marker")
        if "New report file(s):" not in txt and "Report found during wait:" not in txt and "REPORT OK:" not in txt:
            add(findings, "error", latest, "Latest run log does not show a created report file")

        skip_screenshots = "Screenshot regeneration skipped" in txt
        screenshot_success = "Test screenshots regenerated successfully" in txt
        screenshot_timeout = bool(re.search(r"Screenshot regeneration timed out\s*\(>120s\)", txt))
        screenshot_failed = "Screenshot regeneration failed" in txt or "Screenshot regeneration error" in txt
        if screenshot_timeout:
            add(findings, "error", latest, "Screenshot regeneration timed out (>120s); Excel killed - check/read logs, fix cause, rerun", None)
        elif screenshot_failed:
            add(findings, "error", latest, "Screenshot regeneration failed; check logs and fix", None)
        elif skip_screenshots:
            add(findings, "warn", latest, "Screenshot regeneration was skipped (--no-screenshots); screenshots are required")
        elif screenshot_success:
            pass
        elif "Screenshot regeneration" not in txt:
            add(findings, "warn", latest, "Latest run log does not mention screenshot regeneration result")

        # check expected report screenshot files exist and are not stale vs latest report
        try:
            rp = find_latest_report(None)
            if rp and DOC_SHOTS.exists():
                r_m = rp.stat().st_mtime
                for shot in ("report-dashboard.png", "report-raw-data.png"):
                    sp = DOC_SHOTS / shot
                    if not sp.exists():
                        add(findings, "warn", sp, f"Expected screenshot missing: {shot}")
                    elif sp.stat().st_mtime + 5 < r_m:
                        if skip_screenshots or screenshot_timeout or screenshot_failed:
                            add(findings, "warn", sp, f"{shot} is older than latest report (screenshots not completed successfully)")
                        else:
                            add(findings, "warn", sp, f"{shot} is older than latest report (screenshots may be stale)")
        except Exception:
            pass


def find_latest_report(explicit=None):
    if explicit:
        p = Path(explicit)
        if not p.is_absolute():
            p = ROOT / p
        return p if p.exists() else None
    if not REPORTS_DIR.exists():
        return None
    cands = [p for p in REPORTS_DIR.iterdir() if p.is_file() and p.suffix.lower() == ".xlsx" and not p.name.startswith("~$")]
    if not cands:
        return None
    return max(cands, key=lambda p: p.stat().st_mtime)


def scan_report(findings, report_path):
    if report_path is None:
        add(findings, "error", REPORTS_DIR, "No report .xlsx found")
        return None
    if not report_path.exists():
        add(findings, "error", report_path, "Report file does not exist")
        return None
    if report_path.name.startswith("~$"):
        add(findings, "error", report_path, "Selected file is an Excel lock/temp file, not a report")

    try:
        size = report_path.stat().st_size
        if size < 10_000:
            add(findings, "warn", report_path, f"Report file is unexpectedly small ({size} bytes)")
    except Exception as e:
        add(findings, "warn", report_path, f"Could not stat report: {e}")

    if openpyxl is None:
        add(findings, "warn", report_path, "openpyxl not available; skipped workbook structure checks")
        return report_path

    try:
        wb = openpyxl.load_workbook(report_path, read_only=False, data_only=True)
    except Exception as e:
        add(findings, "error", report_path, f"Could not open workbook: {e}")
        return report_path

    try:
        names = set(wb.sheetnames)
        missing = sorted(REQUIRED_SHEETS - names)
        if missing:
            add(findings, "error", report_path, f"Missing required sheet(s): {', '.join(missing)}", extra={"sheets": wb.sheetnames})
        standard = REQUIRED_SHEETS | OPTIONAL_SHEETS
        manager_like = [n for n in wb.sheetnames if n not in standard]
        if not manager_like:
            add(findings, "warn", report_path, "No manager/data tabs found besides Dashboard/Raw Data/Audit Log", extra={"sheets": wb.sheetnames})

        if "Raw Data" in names:
            ws = wb["Raw Data"]
            headers = []
            for c in range(1, min(60, (ws.max_column or 0) + 1)):
                v = ws.cell(1, c).value
                if v is None or str(v).strip() == "":
                    continue
                headers.append(str(v).strip())
            hset = set(headers)
            miss_h = sorted(RAW_EXPECTED_HEADERS_SUBSET - hset)
            if miss_h:
                add(findings, "warn", report_path, f"Raw Data missing expected header(s): {', '.join(miss_h)}", extra={"headers": headers[:40]})
            if (ws.max_row or 0) < 2:
                add(findings, "warn", report_path, "Raw Data appears empty (no data rows)")

        if "Dashboard" in names:
            ws = wb["Dashboard"]
            found_title = False
            for r in range(1, 6):
                for c in range(1, 12):
                    v = ws.cell(r, c).value
                    if v and "PENDING PICKUP" in str(v).upper():
                        found_title = True
                        break
                if found_title:
                    break
            if not found_title:
                add(findings, "warn", report_path, "Dashboard title text not found near top-left")

        if "Audit Log" in names:
            ws = wb["Audit Log"]
            txts = []
            for r in range(1, 30):
                for c in range(1, 4):
                    v = ws.cell(r, c).value
                    if v:
                        txts.append(str(v))
            blob = "\n".join(txts)
            for needle in ("Report Generated", "Source File", "PENDING PICKUP REPORT - AUDIT LOG"):
                if needle.lower() not in blob.lower():
                    add(findings, "warn", report_path, f"Audit Log missing expected text: {needle}")
        else:
            add(findings, "warn", report_path, "Audit Log sheet not found")
    finally:
        try:
            wb.close()
        except Exception:
            pass

    return report_path


def scan_artifacts(findings, report_path):
    if REPORTS_DIR.exists():
        locks = [p for p in REPORTS_DIR.iterdir() if p.name.startswith("~$")]
        for p in locks:
            age = datetime.now().timestamp() - p.stat().st_mtime
            sev = "warn" if age < 3600 else "info"
            add(findings, sev, p, "Excel lock/temp file present in Reports (may be stale if Excel not open)")

    marker = LOG_DIR / "pending-pickup-complete.txt"
    if marker.exists():
        try:
            mtxt = marker.read_text(encoding="utf-8", errors="replace")
            m = re.search(r"saved=(.+)", mtxt)
            if m:
                saved = Path(m.group(1).strip())
                if not saved.exists():
                    add(findings, "warn", marker, "Completion marker references missing report file", extra={"saved": str(saved)})
                elif report_path and saved.resolve() != report_path.resolve():
                    add(findings, "warn", marker, "Completion marker does not point to the newest report", extra={"saved": str(saved), "latest": str(report_path)})
        except Exception as e:
            add(findings, "warn", marker, f"Could not read completion marker: {e}")

    expected_shots = [
        DOC_SHOTS / "report-dashboard.png",
        DOC_SHOTS / "report-raw-data.png",
        DOC_SHOTS / "template-start-here.png",
        DOC_SHOTS / "template-manager-email.png",
        DOC_SHOTS / "template-director-email.png",
    ]
    for p in expected_shots:
        if not p.exists():
            add(findings, "warn", p, "Expected screenshot/artifact missing")
        else:
            try:
                if p.stat().st_size < 1024:
                    add(findings, "warn", p, f"Screenshot is very small ({p.stat().st_size} bytes)")
            except Exception:
                pass

    if report_path and DOC_SHOTS.exists():
        r_m = report_path.stat().st_mtime
        skip_screenshots = False
        screenshot_problem = False
        runs = sorted(LOG_DIR.glob("pending-pickup_*.log"), key=lambda p: p.stat().st_mtime) if LOG_DIR.exists() else []
        if runs:
            try:
                ltxt = runs[-1].read_text(encoding="utf-8", errors="replace")
                skip_screenshots = "Screenshot regeneration skipped" in ltxt
                screenshot_problem = bool(re.search(r"Screenshot regeneration timed out\s*\(>120s\)", ltxt)) or ("Screenshot regeneration failed" in ltxt) or ("Screenshot regeneration error" in ltxt)
            except Exception:
                pass
        dash = DOC_SHOTS / "report-dashboard.png"
        if dash.exists() and dash.stat().st_mtime + 5 < r_m:
            if skip_screenshots or screenshot_problem:
                add(findings, "warn", dash, "report-dashboard.png is older than latest report (screenshots not completed successfully)")
            else:
                add(findings, "warn", dash, "report-dashboard.png is older than latest report (screenshots may be stale)")


def summarize(findings):
    counts = {"error": 0, "warn": 0, "info": 0}
    for f in findings:
        counts[f["severity"]] = counts.get(f["severity"], 0) + 1
    return counts


def main():
    ap = argparse.ArgumentParser(description="Scan Pending Pickup logs and report artifacts for issues")
    ap.add_argument("--report", default=None, help="Path to report .xlsx (default: newest in Reports/)")
    ap.add_argument("--latest", action="store_true", help="Scan only latest run-related logs (faster)")
    ap.add_argument("--json", action="store_true", help="Output JSON")
    ap.add_argument("--strict", action="store_true", help="Exit 1 on warnings as well as errors")
    args = ap.parse_args()

    findings = []
    scan_logs(findings, latest_only=args.latest)
    report_path = find_latest_report(args.report)
    scan_report(findings, report_path)
    scan_artifacts(findings, report_path)

    counts = summarize(findings)
    result = {
        "ok": counts.get("error", 0) == 0 and (not args.strict or counts.get("warn", 0) == 0),
        "report": str(report_path) if report_path else None,
        "counts": counts,
        "findings": findings,
    }

    if args.json:
        print(json.dumps(result, indent=2))
    else:
        print(f"Report: {report_path if report_path else '(none)'}")
        print(f"Findings: {counts.get('error',0)} error(s), {counts.get('warn',0)} warning(s), {counts.get('info',0)} info")
        for f in findings:
            loc = f"{f['source']}" + (f":{f['line']}" if f.get("line") else "")
            print(f"[{f['severity'].upper()}] {loc} - {f['message']}")
            txt = (f.get("extra") or {}).get("text")
            if txt:
                print(f"  > {txt}")
        if result["ok"]:
            print("OK: no blocking issues detected." if not args.strict else "OK: no issues detected (strict).")
        else:
            print("ISSUES DETECTED.")

    sys.exit(0 if result["ok"] else 1)


if __name__ == "__main__":
    main()
