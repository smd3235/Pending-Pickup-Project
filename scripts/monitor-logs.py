#!/usr/bin/env python3
"""Monitor Pending Pickup logs for errors, stalls, and run health.

Scans/tails pending-pickup run logs, macro-debug, watcher, and answer-dialog logs.

Usage:
  python scripts/monitor-logs.py                 # summarize newest run log
  python scripts/monitor-logs.py --follow        # tail newest run log live (Ctrl+C to stop)
  python scripts/monitor-logs.py --all           # summarize newest run + recent macro-debug issues
  python scripts/monitor-logs.py --log <path>    # specific log file
  python scripts/monitor-logs.py --stall-seconds 15
"""
from __future__ import annotations

import argparse
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
LOG_DIR = ROOT / "log"

TS_RE = re.compile(r"^(\d{4}-\d{2}-\d{2} \d{1,2}:\d{2}:\d{2})\s*\|\s*(.*)$")
ERR_PATTERNS = [
    (re.compile(r"\bFAILED\b|\bKILLED\b|\bERROR\b|VBA ERROR DIALOG|\bErr=\s*[1-9]", re.I), "ERROR"),
    (re.compile(r"Traceback \(most recent call last\)|pywintypes\.com_error|Exception occurred", re.I), "ERROR"),
    (re.compile(r"RPC server is unavailable|remote procedure call failed", re.I), "ERROR"),
    (re.compile(r"timeout|timed out|exceeded .* without a new report file", re.I), "ERROR"),
    (re.compile(r"Completion marker was not written", re.I), "ERROR"),
    (re.compile(r"\bWARNING\b", re.I), "WARN"),
]
WARN_LINE_PATTERNS = [
    (re.compile(r"~\$Pending Pickup.*REPORT OK|~\$Pending Pickup.*New report file", re.I), "WARN", "lock/temp file treated as report"),
    (re.compile(r"Could not set VBAWarnings|Could not set trusted location", re.I), "WARN", "registry/setup warning"),
    (re.compile(r"Email/template sync skipped", re.I), "WARN", "sync skipped"),
    (re.compile(r"Snapshot script not found|Screenshot regeneration failed|Screenshot regeneration error", re.I), "WARN", "screenshot issue"),
]
BENIGN = [
    re.compile(r"Run call returned without COM error", re.I),
    re.compile(r"Email/template sync check: email config matches no-macro", re.I),
    re.compile(r"MISSING PICKUP DATE.*label=", re.I),
]


def parse_ts(line: str):
    m = TS_RE.match(line.rstrip("\n\r"))
    if not m:
        return None, line.rstrip("\n\r")
    ts_s, rest = m.group(1), m.group(2)
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(ts_s, fmt), rest
        except Exception:
            pass
    # tolerate single-digit hour (seen in some VBA logs like 2026-11-17 12:7:14)
    try:
        parts = ts_s.replace("-", " ").replace(":", " ").split()
        if len(parts) == 6:
            y, mo, d, h, mi, s = map(int, parts)
            return datetime(y, mo, d, h, mi, s), rest
    except Exception:
        pass
    return None, line.rstrip("\n\r")


def classify(line: str):
    for b in BENIGN:
        if b.search(line):
            return None, None
    for pat, label, msg in WARN_LINE_PATTERNS:
        if pat.search(line):
            return label, msg
    for pat, label in ERR_PATTERNS:
        if pat.search(line):
            return label, label.lower()
    return None, None


def newest_run_log():
    if not LOG_DIR.exists():
        return None
    logs = sorted(LOG_DIR.glob("pending-pickup_*.log"), key=lambda p: p.stat().st_mtime)
    return logs[-1] if logs else None


def read_lines(path: Path):
    try:
        return path.read_text(encoding="utf-8", errors="replace").splitlines()
    except Exception:
        return []


def tail_recent(path: Path, n=20):
    lines = read_lines(path)
    return lines[-n:]


def summarize_file(path: Path, stall_seconds: int, show_lines=True):
    lines = read_lines(path)
    errors = 0
    warns = 0
    stalls = 0
    findings = []
    prev_dt = None
    prev_raw = None
    report_ok_count = 0
    for i, raw in enumerate(lines, 1):
        if "REPORT OK:" in raw:
            report_ok_count += 1
        sev, why = classify(raw)
        if sev == "ERROR":
            errors += 1
            findings.append((i, "ERROR", why or "error", raw.strip()))
        elif sev == "WARN":
            warns += 1
            findings.append((i, "WARN", why or "warning", raw.strip()))
        dt, _ = parse_ts(raw)
        if dt and prev_dt:
            delta = (dt - prev_dt).total_seconds()
            if delta >= stall_seconds:
                stalls += 1
                findings.append((i, "STALL", f"+{int(delta)}s gap", (prev_raw or "").strip()[:200]))
        if dt:
            prev_dt = dt
            prev_raw = raw
    if report_ok_count > 3:
        warns += 1
        findings.append((None, "WARN", f"repeated REPORT OK lines ({report_ok_count})", "watchdog may be spamming"))

    outcome = None
    for raw in reversed(lines[-80:] if len(lines) > 80 else lines):
        if "Final outcome:" in raw or "run finished (outcome=" in raw:
            outcome = raw.strip()
            break
    has_marker = any("Completion marker" in l and "not written" not in l for l in lines[-120:])
    has_report = any("New report file" in l or "Report found during wait" in l or "REPORT OK:" in l for l in lines)

    print(f"Log: {path}")
    print(f"  lines={len(lines)} errors={errors} warnings={warns} stalls(>={stall_seconds}s)={stalls}")
    if outcome:
        print(f"  outcome: {outcome}")
    print(f"  completion_marker_confirmed={has_marker} report_created_logged={has_report}")
    if lines:
        print(f"  first: {lines[0]}")
        print(f"  last : {lines[-1]}")
    if show_lines and findings:
        print("  findings:")
        for ln, sev, why, text in findings[-40:]:
            loc = f":{ln}" if ln else ""
            print(f"    [{sev}] {why}{loc}: {text[:300]}")

    # companion logs
    for name, label in [("watcher.log", "watcher"), ("answer-dialogs.log", "answer-dialogs"), ("compile-test.log", "compile-test")]:
        p = LOG_DIR / name
        if p.exists():
            tail = tail_recent(p, 8)
            if tail:
                # show dialog/error-ish lines preferentially
                interesting = [t for t in tail if any(k in t for k in ("DIALOG", "ERROR", "VBA", "FAILED", "KILLED", "RESULT"))]
                show = interesting[-5:] if interesting else tail[-3:]
                print(f"  {label} tail ({p.name}):")
                for t in show:
                    print(f"    {t}")
    return errors, warns, stalls


def follow_file(path: Path, stall_seconds: int):
    print(f"Following {path} (stall threshold {stall_seconds}s). Ctrl+C to stop.")
    prev_dt = None
    try:
        with path.open("r", encoding="utf-8", errors="replace") as f:
            f.seek(0, os.SEEK_END)
            while True:
                line = f.readline()
                if not line:
                    time.sleep(0.4)
                    continue
                raw = line.rstrip("\n\r")
                sev, why = classify(raw)
                dt, _ = parse_ts(raw)
                tags = []
                if sev:
                    tags.append(f"[{sev}:{why}]")
                if dt and prev_dt:
                    delta = (dt - prev_dt).total_seconds()
                    if delta >= stall_seconds:
                        tags.append(f"[STALL +{int(delta)}s]")
                if tags:
                    print("ALERT " + " ".join(tags) + " " + raw)
                else:
                    print(raw)
                if dt:
                    prev_dt = dt
    except KeyboardInterrupt:
        print("\nStopped.")


def main():
    ap = argparse.ArgumentParser(description="Monitor Pending Pickup logs for errors and stalls")
    ap.add_argument("--follow", "-f", action="store_true", help="tail newest run log live")
    ap.add_argument("--log", default=None, help="specific log file to monitor/summarize")
    ap.add_argument("--stall-seconds", type=int, default=10, help="gap threshold to flag as stall (default 10)")
    ap.add_argument("--all", action="store_true", help="also summarize recent macro-debug issues")
    args = ap.parse_args()

    if args.log:
        path = Path(args.log)
        if not path.is_absolute():
            path = ROOT / path
    else:
        path = newest_run_log()

    if not path or not path.exists():
        print(f"No run log found in {LOG_DIR}")
        return 1

    if args.follow:
        follow_file(path, args.stall_seconds)
        return 0

    e, w, s = summarize_file(path, args.stall_seconds)

    if args.all:
        mdbg = LOG_DIR / "macro-debug.log"
        if mdbg.exists():
            print("\nmacro-debug.log (recent issues):")
            lines = read_lines(mdbg)[-200:]
            count = 0
            for i, raw in enumerate(lines, 1):
                sev, why = classify(raw)
                if sev:
                    print(f"  [{sev}] {why}: {raw.strip()[:300]}")
                    count += 1
            if count == 0:
                print("  (no flagged issues in recent tail)")

    return 1 if e else 0


if __name__ == "__main__":
    sys.exit(main())
