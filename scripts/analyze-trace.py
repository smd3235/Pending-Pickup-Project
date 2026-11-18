#!/usr/bin/env python3
"""Analyze VBA per-run logs for milestones, TRACE pairing, gaps, and anomalies.

Targets Pending Pickup VBA run logs (files containing
"=== CreatePendingPickupReport started ===" with run metadata).

Usage:
  python scripts/analyze-trace.py                 # latest VBA per-run log
  python scripts/analyze-trace.py --log <path>    # specific log
  python scripts/analyze-trace.py --all           # analyze all matching logs
  python scripts/analyze-trace.py --stall-seconds 20
  python scripts/analyze-trace.py --json
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
LOG_DIR = ROOT / "log"

TS_RE = re.compile(r"^(\d{4}-\d{2}-\d{2} \d{1,2}:\d{1,2}:\d{2})\s*\|\s*(.*)$")

REQUIRED_MILESTONES = [
    ("run started", re.compile(r"=== CreatePendingPickupReport started ==="), "error"),
    ("source selected", re.compile(r"CURRENT selected:"), "warn"),
    ("source dimensions/counts", re.compile(r"currentCount=\d+ managers=\d+"), "warn"),
    ("aging stats", re.compile(r"Aging: 30=\d+ 60=\d+ 90=\d+ missing=\d+", re.I), "warn"),
    ("dashboard built", re.compile(r"BuildDashboard done|BuildDashboard done - all cards and table created"), "warn"),
    ("raw data created", re.compile(r"Raw Data sheet created"), "warn"),
    ("save path", re.compile(r"Saving to:|TRACE: BuildOutputPath done"), "warn"),
    ("save done", re.compile(r"\bSave done\b|TRACE: SaveAs done"), "error"),
    ("completion marker", re.compile(r"Completion marker written:"), "error"),
    ("run completed", re.compile(r"=== CreatePendingPickupReport completed ==="), "error"),
]

# (start_re, done_re, label, opts)
# opts: keyed -> capture group "name" identifies instance; allow_done_only -> done without start is ok
TRACE_SPECS = [
    (re.compile(r"NormalizeDateColumn start\b"), re.compile(r"NormalizeDateColumn done\b"), "NormalizeDateColumn", {}),
    (re.compile(r"Value2 start\b"), re.compile(r"Value2 done\b"), "Value2", {}),
    (re.compile(r"TRACE: ApplySourceDateFormats start\b"), re.compile(r"TRACE: ApplySourceDateFormats done\b"), "ApplySourceDateFormats", {"allow_done_only": True}),
    (re.compile(r"TRACE: Workbooks\.Add start\b"), re.compile(r"TRACE: Workbooks\.Add done\b|Output workbook created\b"), "Workbooks.Add", {"allow_done_only": True}),
    (re.compile(r"FormatManagerSheet start: (?P<name>.+?) rows="), re.compile(r"FormatManagerSheet done: (?P<name>.+)$"), "FormatManagerSheet", {"keyed": True}),
    (re.compile(r"TRACE: BuildOutputArray start\b"), re.compile(r"TRACE: BuildOutputArray done\b"), "BuildOutputArray", {}),
    (re.compile(r"TRACE: writing raw data values to sheet\b"), re.compile(r"TRACE: raw data values written\b"), "raw data write", {}),
    (re.compile(r"TRACE: FormatRawData start\b"), re.compile(r"TRACE: FormatRawData done\b"), "FormatRawData", {}),
    (re.compile(r"TRACE: CopyColumnFormats start\b"), re.compile(r"TRACE: CopyColumnFormats done\b"), "CopyColumnFormats", {}),
    (re.compile(r"TRACE: MatchSourceColumnWidths start\b"), re.compile(r"TRACE: MatchSourceColumnWidths done\b"), "MatchSourceColumnWidths", {}),
    (re.compile(r"TRACE: ApplyDateFormatsToOutput start\b"), re.compile(r"TRACE: ApplyDateFormatsToOutput done\b"), "ApplyDateFormatsToOutput", {}),
    (re.compile(r"TRACE: ApplyDateFormatByHeader \((?P<name>.+)\) start\b"), re.compile(r"TRACE: ApplyDateFormatByHeader \((?P<name>.+)\) done\b"), "ApplyDateFormatByHeader", {"keyed": True}),
    (re.compile(r"TRACE: AuditWriteSheet start\b"), re.compile(r"TRACE: AuditWriteSheet done\b|Audit sheet written\b"), "AuditWriteSheet", {"allow_done_only": True}),
    (re.compile(r"TRACE: AuditLogEvent SAVE start\b"), re.compile(r"TRACE: AuditLogEvent SAVE done\b"), "AuditLogEvent SAVE", {}),
    (re.compile(r"TRACE: SaveAs start\b"), re.compile(r"TRACE: SaveAs done\b"), "SaveAs", {}),
    (re.compile(r"TRACE: HideUnusedCells start\b"), re.compile(r"TRACE: HideUnusedCells done\b"), "HideUnusedCells", {"allow_done_only": True}),
    (re.compile(r"TRACE: BuildOutputPath start\b"), re.compile(r"TRACE: BuildOutputPath done\b"), "BuildOutputPath", {"allow_done_only": True}),
]

ISSUE_PATTERNS = [
    (re.compile(r"\bFAILED\b|\bErr=\s*[1-9]\b|\bERROR\b", re.I), "error"),
    (re.compile(r"\btraceback\b|\bexception\b", re.I), "error"),
    (re.compile(r"\btimeout\b|timed out", re.I), "error"),
    (re.compile(r"RPC server is unavailable|remote procedure call failed", re.I), "error"),
    (re.compile(r"could not|unavailable", re.I), "warn"),
]


def parse_ts(s: str):
    try:
        parts = re.split(r"[- :]", s.strip())
        if len(parts) == 6:
            y, mo, d, h, mi, sec = map(int, parts)
            return datetime(y, mo, d, h, mi, sec)
    except Exception:
        pass
    return None


def is_vba_run_log(path: Path):
    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except Exception:
        return False
    # VBA per-run trace created by modCore/modLogging, not the Python runner log
    return (
        "=== CreatePendingPickupReport started ===" in text
        and "Run log:" in text
        and "Completion marker:" in text
        and "VBAWarnings registry set" not in text
    )


def find_logs(specific=None, all_logs=False):
    if specific:
        p = Path(specific)
        if not p.is_absolute():
            p = ROOT / p
        return [p] if p.exists() else []
    if not LOG_DIR.exists():
        return []
    cands = [p for p in LOG_DIR.glob("pending-pickup_*.log") if is_vba_run_log(p)]
    cands.sort(key=lambda p: p.stat().st_mtime)
    if all_logs:
        return cands
    return cands[-1:] if cands else []


def analyze(path: Path, stall_seconds: int):
    text = path.read_text(encoding="utf-8", errors="replace")
    lines = text.splitlines()
    findings = []
    counts = {"error": 0, "warn": 0, "info": 0}

    def add(sev, msg, line=None, extra=None):
        findings.append({"severity": sev, "line": line, "message": msg, "extra": extra or {}})
        counts[sev] = counts.get(sev, 0) + 1

    # milestones
    for label, pat, sev in REQUIRED_MILESTONES:
        if not pat.search(text):
            add(sev, f"Missing milestone: {label}")

    if "=== CreatePendingPickupReport completed ===" in text and "Completion marker written:" not in text:
        add("error", "Run marked completed but completion marker not written")

    # issue patterns
    for i, raw in enumerate(lines, 1):
        if "MISSING PICKUP DATE" in raw and "label=" in raw:
            continue
        for pat, sev in ISSUE_PATTERNS:
            if pat.search(raw):
                add(sev, "Flagged pattern in log", i, {"text": raw.strip()[:300]})
                break

    # gaps
    prev_dt = None
    prev_i = None
    prev_raw = None
    for i, raw in enumerate(lines, 1):
        m = TS_RE.match(raw)
        if not m:
            continue
        dt = parse_ts(m.group(1))
        if dt and prev_dt:
            delta = (dt - prev_dt).total_seconds()
            if delta >= stall_seconds:
                add("warn", f"Large gap +{int(delta)}s between log lines", i, {"from_line": prev_i, "from": (prev_raw or "")[:200], "to": raw.strip()[:200]})
        if dt:
            prev_dt, prev_i, prev_raw = dt, i, raw.strip()

    # TRACE pairing
    open_starts = {}  # (label,key) -> count
    bd_starts = 0
    bd_dones = 0
    bd_start_re = re.compile(r"BuildDashboard start\b")
    bd_done_re = re.compile(r"BuildDashboard done\b")

    for i, raw in enumerate(lines, 1):
        m = TS_RE.match(raw)
        body = m.group(2) if m else raw

        # special: BuildDashboard has nested/internal calls and different done strings
        if bd_start_re.search(body) and "BuildDashboard start" == body.strip():
            bd_starts += 1
            continue
        if bd_done_re.search(body) and body.strip().startswith("BuildDashboard done"):
            bd_dones += 1
            continue

        for sp, dp, label, opts in TRACE_SPECS:
            sm = sp.search(body)
            if sm:
                keyname = sm.groupdict().get("name") if opts.get("keyed") else None
                k = (label, keyname)
                open_starts[k] = open_starts.get(k, 0) + 1
                break
            dm = dp.search(body)
            if dm:
                keyname = dm.groupdict().get("name") if opts.get("keyed") else None
                k = (label, keyname)
                if open_starts.get(k, 0) > 0:
                    open_starts[k] -= 1
                else:
                    if not opts.get("allow_done_only"):
                        add("warn", f"TRACE done without matching start: {label}" + (f" ({keyname})" if keyname else ""), i, {"text": body.strip()[:200]})
                break

    if bd_starts or bd_dones:
        if bd_starts == 0 and bd_dones > 0:
            add("warn", f"BuildDashboard done without start (dones={bd_dones})")
        elif bd_starts > 0 and bd_dones == 0:
            add("error", f"BuildDashboard started but never completed (starts={bd_starts})")
        elif bd_starts != bd_dones:
            add("warn", f"BuildDashboard start/done count mismatch: starts={bd_starts} dones={bd_dones}")

    for (label, keyname), cnt in open_starts.items():
        if cnt > 0:
            sev = "error" if label in ("SaveAs", "AuditWriteSheet") else "warn"
            add(sev, f"TRACE start without done: {label}" + (f" ({keyname})" if keyname else "") + f" (open={cnt})")

    # manager tab sanity
    created = len(re.findall(r"Creating manager tab:", text))
    done_tabs = len(re.findall(r"\smanager tab done:", text))
    if created or done_tabs:
        if created != done_tabs:
            add("warn", f"Manager tab start/done mismatch: created={created} done={done_tabs}")

    out_path_m = re.search(r"Output file: (.+)", text)
    out_path = out_path_m.group(1).strip() if out_path_m else None
    if out_path:
        p = Path(out_path)
        if not p.exists():
            add("warn", "Output file path logged but file not found on disk", extra={"path": out_path})

    summary = {
        "file": str(path),
        "lines": len(lines),
        "manager_tabs_created": created,
        "manager_tabs_done": done_tabs,
        "output_file": out_path,
    }
    m = re.search(r"currentCount=(\d+) managers=(\d+)", text)
    if m:
        summary["currentCount"] = int(m.group(1))
        summary["managers"] = int(m.group(2))
    m = re.search(r"Aging: 30=(\d+) 60=(\d+) 90=(\d+) missing=(\d+)", text)
    if m:
        summary["aging"] = {"30": int(m.group(1)), "60": int(m.group(2)), "90": int(m.group(3)), "missing": int(m.group(4))}
    m = re.search(r"Run duration \(seconds\):\s*([0-9.]+)", text)
    if m:
        summary["run_duration_seconds"] = float(m.group(1))
    summary["counts"] = counts
    summary["findings"] = findings
    return summary


def main():
    ap = argparse.ArgumentParser(description="Analyze VBA per-run trace logs for gaps and missing milestones")
    ap.add_argument("--log", default=None)
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--stall-seconds", type=int, default=15)
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()

    logs = find_logs(args.log, args.all)
    if not logs:
        msg = {"ok": False, "error": "No VBA per-run logs found"}
        if args.json:
            print(json.dumps(msg, indent=2))
        else:
            print("No VBA per-run logs found.")
        return 1

    results = [analyze(p, args.stall_seconds) for p in logs]
    ok = all(r["counts"].get("error", 0) == 0 for r in results)

    if args.json:
        print(json.dumps({"ok": ok, "results": results}, indent=2))
    else:
        for r in results:
            print(f"Trace: {r['file']}")
            c = r["counts"]
            print(f"  lines={r['lines']} errors={c.get('error',0)} warnings={c.get('warn',0)} info={c.get('info',0)} managers={r.get('manager_tabs_created',0)}/{r.get('manager_tabs_done',0)}")
            if "currentCount" in r:
                print(f"  currentCount={r['currentCount']} managers={r.get('managers')} aging={r.get('aging')} duration_s={r.get('run_duration_seconds')}")
            if r.get("output_file"):
                print(f"  output_file={r['output_file']}")
            for f in r["findings"]:
                loc = f":{f['line']}" if f.get("line") else ""
                print(f"  [{f['severity'].upper()}]{loc} {f['message']}")
                ex = f.get("extra") or {}
                if ex.get("text"):
                    print(f"    > {ex['text']}")
            if not r["findings"]:
                print("  OK: no issues detected in trace.")
            print()
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
