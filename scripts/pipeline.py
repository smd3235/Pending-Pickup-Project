#!/usr/bin/env python3
"""Pending Pickup pipeline: compile -> sync -> run -> screenshots -> scan/validate.

Hard rule: non-screenshot steps must complete within 30 seconds; on timeout or
failure the step is stopped, Excel is killed, relevant logs are printed, and the
pipeline exits non-zero so the issue can be fixed before retrying.

Screenshots/exports are required and must be allowed to finish. They may take
longer than 30 seconds, but if a screenshot/export exceeds 120 seconds (2 minutes)
treat it as a failure/hang: kill Excel/app, print logs, and exit non-zero.
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

from excel_guard import DEFAULT_TIMEOUT, dump_recent_logs, kill_excel_processes, log_pipeline

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
PYTHON = sys.executable
PIPELINE_LOG = ROOT / "log" / "pipeline.log"
SCREENSHOT_TIMEOUT = 120  # seconds - screenshots may take >30s but not >2m


def _ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def log(msg: str) -> None:
    line = f"{_ts()} | {msg}"
    print(line)
    PIPELINE_LOG.parent.mkdir(parents=True, exist_ok=True)
    with PIPELINE_LOG.open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def print_log_tails(extra: list[str] | None = None) -> None:
    names = ["pipeline.log", "excel-hangs.log", "watcher.log", "compile-test.log"]
    if extra:
        names = extra + names
    print("\n=== Recent logs (for diagnosis) ===")
    txt = dump_recent_logs(names, max_lines=50)
    if txt.strip():
        print(txt)
    else:
        print("(no log tails available)")


def _ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def log(msg: str) -> None:
    line = f"{_ts()} | {msg}"
    print(line)
    PIPELINE_LOG.parent.mkdir(parents=True, exist_ok=True)
    with PIPELINE_LOG.open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def print_log_tails(extra: list[str] | None = None) -> None:
    names = ["pipeline.log", "excel-hangs.log", "watcher.log", "compile-test.log"]
    if extra:
        names = extra + names
    print("\n=== Recent logs (for diagnosis) ===")
    txt = dump_recent_logs(names, max_lines=50)
    if txt.strip():
        print(txt)
    else:
        print("(no log tails available)")


def run_step(
    label: str,
    cmd: list[str],
    timeout: int | None = DEFAULT_TIMEOUT,
    env: dict | None = None,
    *,
    kill_excel: bool = True,
    screenshot_step: bool = False,
) -> None:
    """Run a pipeline step.

    Non-screenshot steps: 30s timeout, kill Excel on timeout/failure.
    Screenshot/export steps: 120s timeout, allowed to finish under 2 minutes,
    but Excel is killed if the 120s timeout is exceeded or on failure.
    """
    timeout_desc = f"{timeout}s" if timeout is not None else "none"
    step_kind = "screenshot" if screenshot_step else "standard"
    print(f"\n=== STEP: {label} ===")
    print(" ".join(cmd))
    log(f"STEP {label}: STARTED (timeout={timeout_desc}, kill_excel={kill_excel}, kind={step_kind})")
    start = time.time()
    try:
        proc = subprocess.run(
            cmd,
            cwd=str(ROOT),
            capture_output=True,
            text=True,
            timeout=timeout,
            env=env or os.environ.copy(),
        )
        dur = time.time() - start
        if proc.stdout:
            print(proc.stdout.rstrip())
        if proc.stderr:
            print(proc.stderr.rstrip(), file=sys.stderr)
        if proc.returncode != 0:
            log(f"STEP {label}: FAILED exit={proc.returncode} dur={dur:.2f}s")
            print(f"FAILED: {label} (exit {proc.returncode}, {dur:.2f}s)", file=sys.stderr)
            if kill_excel:
                kill_excel_processes(reason=f"step failed: {label}")
            print_log_tails()
            sys.exit(proc.returncode or 1)
        log(f"STEP {label}: OK dur={dur:.2f}s")
        print(f"OK: {label} ({dur:.2f}s)")
    except subprocess.TimeoutExpired as e:
        dur = time.time() - start
        log(f"STEP {label}: TIMEOUT after {timeout}s (dur={dur:.2f}s)")
        if screenshot_step:
            print(f"TIMEOUT: {label} exceeded {timeout}s (screenshot limit). Killing Excel and stopping - check logs, fix, rerun.", file=sys.stderr)
        else:
            print(f"TIMEOUT: {label} exceeded {timeout}s. Killing Excel and stopping.", file=sys.stderr)
        try:
            if e.stdout:
                out = e.stdout if isinstance(e.stdout, str) else e.stdout.decode("utf-8", "replace")
                print(out[-4000:])
            if e.stderr:
                err = e.stderr if isinstance(e.stderr, str) else e.stderr.decode("utf-8", "replace")
                print(err[-4000:], file=sys.stderr)
        except Exception:
            pass
        if kill_excel:
            kill_excel_processes(reason=f"timeout: {label}")
        print_log_tails()
        sys.exit(124)


def main() -> int:
    ap = argparse.ArgumentParser(description="Pending Pickup pipeline (120s non-screenshot, 120s screenshot cap)")
    ap.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT, help=f"Per-step timeout seconds for non-screenshot steps (default {DEFAULT_TIMEOUT})")
    ap.add_argument("--skip-compile", action="store_true")
    ap.add_argument("--skip-sync", action="store_true")
    ap.add_argument("--skip-run", action="store_true")
    ap.add_argument("--skip-scan", action="store_true")
    ap.add_argument("--skip-snapshots", action="store_true", help="Skip screenshot regeneration (debug only; screenshots are required)")
    ap.add_argument("--scan-all-history", action="store_true", help="Scan all historical logs (default scans latest run context)")
    args = ap.parse_args()

    timeout = max(1, int(args.timeout))

    (ROOT / "log").mkdir(parents=True, exist_ok=True)
    log_pipeline("=== PIPELINE START ===")
    log(f"Pipeline start timeout={timeout}s screenshot_timeout={SCREENSHOT_TIMEOUT}s")

    # Fast, no-Excel quality gates first (deps + lint + unit tests + VBA static analysis).
    run_step("deps (pip check)", [PYTHON, "-m", "pip", "check"], timeout=timeout)
    run_step("lint (ruff)", [PYTHON, "-m", "ruff", "check", "--select", "E9,F821,F722", "scripts/", "tests/"], timeout=timeout)
    run_step("unit tests (pytest)", [PYTHON, "-m", "pytest", "tests/", "-q"], timeout=timeout)
    if (SCRIPTS / "vba-analyze.py").exists():
        run_step("vba static analysis (vba-analyze.py)", [PYTHON, str(SCRIPTS / "vba-analyze.py")], timeout=timeout)
    if (SCRIPTS / "screenshot-diff.py").exists():
        run_step(
            "types (mypy)",
            [PYTHON, "-m", "mypy", "--follow-imports=skip",
             "scripts/vba_ocr_helpers.py", "scripts/screenshot-diff.py", "scripts/vba-analyze.py",
             "scripts/pipeline.py", "scripts/test-compile.py", "scripts/vba-error-ocr.py", "tests/"],
            timeout=timeout,
        )

    if not args.skip_compile:
        run_step("compile (test-compile.py)", [PYTHON, str(SCRIPTS / "test-compile.py")], timeout=timeout)

    if not args.skip_sync:
        run_step("sync (sync_template.py)", [PYTHON, str(SCRIPTS / "sync_template.py")], timeout=timeout)

    if not args.skip_run:
        # run-report handles its own 30s VBA/report watchdog; skip post-run screenshot regen
        # so this step stays within the 30s cap (screenshots run as a separate capped step).
        run_step("run-report (run-report.py)", [PYTHON, str(SCRIPTS / "run-report.py"), "--no-screenshots"], timeout=timeout)

    # Screenshots/exports are required: allow to finish, but kill Excel if >120s.
    if args.skip_snapshots:
        log("Screenshot step skipped (--skip-snapshots). Screenshots are required; rerun without skip to complete.")
        print("WARNING: Screenshot step skipped (--skip-snapshots). Screenshots are required.", file=sys.stderr)
    elif not args.skip_run:
        if (SCRIPTS / "snapshot-report.py").exists():
            run_step(
                "snapshot-report",
                [PYTHON, str(SCRIPTS / "snapshot-report.py")],
                timeout=SCREENSHOT_TIMEOUT,
                kill_excel=True,
                screenshot_step=True,
            )
        else:
            log("snapshot-report.py missing; screenshots cannot be generated.")
            print("FAILED: snapshot-report.py missing", file=sys.stderr)
            print_log_tails()
            sys.exit(1)

    # Perceptual-hash regression check on regenerated screenshots.
    if (SCRIPTS / "screenshot-diff.py").exists() and not args.skip_snapshots:
        run_step(
            "screenshot-regression (screenshot-diff.py)",
            [PYTHON, str(SCRIPTS / "screenshot-diff.py")],
            timeout=timeout,
            kill_excel=False,
            screenshot_step=False,
        )

    if not args.skip_scan:
        # Fast, no-Excel scans: keep within 30s each
        run_step("monitor-logs", [PYTHON, str(SCRIPTS / "monitor-logs.py")], timeout=timeout)
        run_step("analyze-trace", [PYTHON, str(SCRIPTS / "analyze-trace.py")], timeout=timeout)
        run_step("validate-report", [PYTHON, str(SCRIPTS / "validate-report.py")], timeout=timeout)
        check_cmd = [PYTHON, str(SCRIPTS / "check-logs.py"), "--latest"]
        if args.scan_all_history:
            check_cmd = [PYTHON, str(SCRIPTS / "check-logs.py")]
        run_step("check-logs", check_cmd, timeout=timeout)
        if (SCRIPTS / "audit-code.py").exists():
            run_step("audit-code", [PYTHON, str(SCRIPTS / "audit-code.py")], timeout=timeout)

    if args.skip_snapshots:
        log_pipeline("=== PIPELINE INCOMPLETE (screenshots skipped) ===")
        log("Pipeline finished checks but screenshots were skipped (required).")
        print("\nPIPELINE INCOMPLETE: screenshots were skipped (required).")
        return 2

    log_pipeline("=== PIPELINE COMPLETE ===")
    log("Pipeline complete: OK")
    print("\nPIPELINE OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
