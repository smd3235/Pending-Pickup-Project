#!/usr/bin/env python3
"""Shared Excel COM guard for Pending Pickup scripts.

Hard rule: nothing should wait longer than DEFAULT_TIMEOUT seconds for Excel/COM
work. If a guarded operation exceeds the timeout, the Excel process is killed,
the hang is logged, and callers should read logs and fix the cause.

Adapted for this project's layout (log/ directory, Windows PowerShell env).
"""
from __future__ import annotations

import os
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

try:
    import psutil
except Exception:  # pragma: no cover
    psutil = None

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent
LOG_DIR = PROJECT_ROOT / "log"
LOG_DIR.mkdir(parents=True, exist_ok=True)

HANG_LOG = LOG_DIR / "excel-hangs.log"
WATCH_LOG = LOG_DIR / "watcher.log"
PIPELINE_LOG = LOG_DIR / "pipeline.log"
WATCHER_SCRIPT = SCRIPT_DIR / "watch-compile-error.py"

DEFAULT_TIMEOUT = 120  # seconds - per-step cap for non-screenshot pipeline steps


def _ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _append(path: Path, msg: str) -> None:
    try:
        with path.open("a", encoding="utf-8") as f:
            f.write(f"{_ts()} | {msg}\n")
    except Exception:
        pass


def log_pipeline(msg: str) -> None:
    line = f"{_ts()} | {msg}"
    print(line)
    _append(PIPELINE_LOG, msg)


def log_hang(msg: str) -> None:
    line = f"{_ts()} | {msg}"
    print(line, file=sys.stderr)
    _append(HANG_LOG, msg)


def log_dialog(msg: str) -> None:
    line = f"{_ts()} | {msg}"
    print(line)
    _append(WATCH_LOG, msg)


def excel_pids() -> set[int]:
    pids: set[int] = set()
    if psutil is None:
        return pids
    try:
        for p in psutil.process_iter(["name", "pid"]):
            try:
                name = (p.info.get("name") or "").lower()
                if name == "excel.exe":
                    pids.add(int(p.info["pid"]))
            except Exception:
                pass
    except Exception:
        pass
    return pids


def kill_excel_processes(reason: str = "forced cleanup") -> list[int]:
    killed: list[int] = []
    if psutil is not None:
        for pid in list(excel_pids()):
            try:
                psutil.Process(pid).kill()
                killed.append(pid)
            except Exception:
                pass
    else:
        try:
            subprocess.run(["taskkill", "/F", "/IM", "excel.exe"], capture_output=True, timeout=10)
        except Exception:
            pass
    if killed:
        log_hang(f"Killed Excel pid(s) {','.join(map(str, killed))} ({reason})")
    return killed


def start_excel(visible: bool = False, display_alerts: bool = False, automation_security: int = 1):
    """Start a COM Excel instance and return (excel, pid).

    PID detection is best-effort via process diff; callers should still guard
    blocking operations with ExcelGuard.
    """
    import win32com.client  # type: ignore

    before = excel_pids()
    excel = win32com.client.DispatchEx("Excel.Application")
    try:
        excel.Visible = bool(visible)
    except Exception:
        pass
    try:
        excel.DisplayAlerts = bool(display_alerts)
    except Exception:
        pass
    try:
        excel.AutomationSecurity = automation_security
    except Exception:
        pass

    pid = None
    deadline = time.time() + 5
    while time.time() < deadline:
        new = excel_pids() - before
        if new:
            pid = sorted(new)[-1]
            break
        time.sleep(0.1)
    if pid:
        log_pipeline(f"Excel started pid={pid} visible={visible}")
    else:
        log_pipeline("Excel started (pid unknown)")
    return excel, pid


class ExcelTimeoutError(TimeoutError):
    """Raised when a guarded Excel operation exceeds the hard timeout."""


class ExcelGuard:
    """Kill Excel if a blocking COM section exceeds timeout seconds."""

    def __init__(self, pid: int | None, timeout: int = DEFAULT_TIMEOUT, label: str = "Excel operation"):
        self.pid = pid
        self.timeout = int(timeout)
        self.label = label
        self._timer: threading.Timer | None = None
        self._fired = threading.Event()

    def _trigger(self) -> None:
        self._fired.set()
        if self.pid:
            try:
                if psutil is not None:
                    psutil.Process(self.pid).kill()
                else:
                    subprocess.run(["taskkill", "/F", "/PID", str(self.pid)], capture_output=True, timeout=10)
            except Exception:
                pass
        else:
            kill_excel_processes(reason=f"timeout>{self.timeout}s pid unknown: {self.label}")
        log_hang(f"TIMEOUT: {self.label} exceeded {self.timeout}s; Excel killed (pid={self.pid})")

    def __enter__(self):
        self._timer = threading.Timer(self.timeout, self._trigger)
        self._timer.daemon = True
        self._timer.start()
        return self

    @property
    def timed_out(self) -> bool:
        return self._fired.is_set()

    def __exit__(self, exc_type, exc, tb):
        if self._timer:
            self._timer.cancel()
        if exc_type is None:
            log_pipeline(f"Excel call completed: {self.label} (pid={self.pid}, timeout={self.timeout}s)")
            return False
        if self._fired.is_set() or isinstance(exc, ExcelTimeoutError):
            log_hang(f"Excel call timed out: {self.label} ({exc})")
        else:
            log_hang(f"Excel call error in {self.label}: {exc}")
        return False


class DialogWatcher:
    """Run watch-compile-error.py alongside Excel work and collect dialog entries."""

    def __init__(self, script: Path | None = None):
        self.script = Path(script) if script else WATCHER_SCRIPT
        self.proc: subprocess.Popen | None = None
        self.start_lines = 0
        self.dialogs: list[str] = []

    def __enter__(self):
        if WATCH_LOG.exists():
            try:
                with WATCH_LOG.open("r", encoding="utf-8", errors="replace") as f:
                    self.start_lines = sum(1 for _ in f)
            except Exception:
                self.start_lines = 0
        else:
            self.start_lines = 0
        if self.script.exists():
            try:
                creationflags = 0
                if hasattr(subprocess, "CREATE_NO_WINDOW"):
                    creationflags = subprocess.CREATE_NO_WINDOW  # type: ignore[attr-defined]
                self.proc = subprocess.Popen(
                    [sys.executable, str(self.script)],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    cwd=str(PROJECT_ROOT),
                    creationflags=creationflags,
                )
                log_pipeline(f"Dialog watcher started (pid={self.proc.pid})")
            except Exception as e:
                log_hang(f"Could not start dialog watcher: {e}")
                self.proc = None
        else:
            log_hang(f"Dialog watcher script missing: {self.script}")
        return self

    def _read_new_dialogs(self) -> list[str]:
        if not WATCH_LOG.exists():
            return []
        try:
            with WATCH_LOG.open("r", encoding="utf-8", errors="replace") as f:
                lines = f.readlines()
        except Exception:
            return []
        out = []
        for raw in lines[self.start_lines:]:
            s = raw.strip()
            if "DIALOG:" in s or "KILLED EXCEL" in s or "dialog" in s.lower():
                out.append(s)
        return out

    def __exit__(self, exc_type, exc, tb):
        collected: list[str] = []
        if self.proc:
            try:
                self.proc.terminate()
                try:
                    self.proc.wait(timeout=3)
                except Exception:
                    self.proc.kill()
            except Exception:
                pass
            try:
                out, _ = self.proc.communicate(timeout=2)
                if out:
                    for line in out.splitlines():
                        if line.strip():
                            collected.append(line.strip())
            except Exception:
                pass
        collected.extend(self._read_new_dialogs())
        seen = set()
        uniq = []
        for d in collected:
            if d not in seen:
                seen.add(d)
                uniq.append(d)
        self.dialogs = uniq
        for d in uniq:
            log_dialog(f"DIALOG DURING SESSION: {d}")
        log_pipeline(f"Dialog watcher stopped ({len(uniq)} dialog(s))")
        return False


def tail_file(path: Path, max_lines: int = 40) -> list[str]:
    try:
        with path.open("r", encoding="utf-8", errors="replace") as f:
            return [ln.rstrip("\n\r") for ln in f.readlines()[-max_lines:]]
    except Exception:
        return []


def dump_recent_logs(log_names: list[str] | None = None, max_lines: int = 40) -> str:
    names = log_names or [
        "pipeline.log",
        "excel-hangs.log",
        "watcher.log",
        "compile-test.log",
    ]
    parts = []
    for name in names:
        p = LOG_DIR / name
        if p.exists():
            lines = tail_file(p, max_lines=max_lines)
            if lines:
                parts.append(f"--- {name} (last {len(lines)}) ---")
                parts.extend(lines)
    # include newest pending-pickup logs
    try:
        runs = sorted(LOG_DIR.glob("pending-pickup_*.log"), key=lambda x: x.stat().st_mtime)[-2:]
        for p in runs:
            lines = tail_file(p, max_lines=max_lines)
            if lines:
                parts.append(f"--- {p.name} (last {len(lines)}) ---")
                parts.extend(lines)
    except Exception:
        pass
    return "\n".join(parts)
