# watch-template-sync.py
# Watches Pending Pickup Template.xlsm for changes and keeps email template
# sections (TO/CC/SUBJECT/BODY on Manager/Director email tabs) in sync with
# Pending Pickup Template (no-macros).xlsx and Project-Local-Build.zip.
#
# Runs python scripts/sync_template.py automatically after a stable save when
# email config differs (or when --always-sync-on-change is set).
# Polling-based (no extra dependencies).
#
# Usage:
#   python scripts/watch-template-sync.py
#   python scripts/watch-template-sync.py --interval 5
#
# Stop with Ctrl+C.

import argparse
import os
import subprocess
import sys
import time
from datetime import datetime

import openpyxl

EMAIL_SHEETS = ["Manager Email Template", "Director Email Template"]
EMAIL_CELLS = ["D5", "D7", "D9", "D11"]  # TO, CC, SUBJECT, BODY


def log(msg, log_path=None):
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"{stamp} | {msg}"
    print(line, flush=True)
    if log_path:
        try:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(line + "\n")
        except Exception:
            pass


def read_email_config(path):
    try:
        wb = openpyxl.load_workbook(path, keep_vba=path.lower().endswith(".xlsm"), data_only=False, read_only=False)
        out = {}
        for name in EMAIL_SHEETS:
            ws = wb[name]
            vals = {}
            for c in EMAIL_CELLS:
                v = ws[c].value
                if v is None:
                    vals[c] = None
                else:
                    # normalize newlines for comparison
                    vals[c] = str(v).replace("\r\n", "\n").replace("\r", "\n")
            out[name] = vals
        try:
            wb.close()
        except Exception:
            pass
        return out, None
    except Exception as e:
        return None, str(e)


def file_stat(path):
    try:
        st = os.stat(path)
        return (st.st_mtime_ns, st.st_size)
    except Exception:
        return None


def is_stable(path, settle_seconds):
    s1 = file_stat(path)
    if s1 is None:
        return False
    time.sleep(settle_seconds)
    s2 = file_stat(path)
    return s1 == s2 and s2 is not None


def main():
    parser = argparse.ArgumentParser(description="Watch template and auto-sync email sections to no-macro/package artifacts")
    parser.add_argument("--interval", type=float, default=4.0, help="Poll interval seconds (default 4)")
    parser.add_argument("--settle", type=float, default=3.0, help="Seconds file must be unchanged before acting (default 3)")
    parser.add_argument("--always-sync-on-change", action="store_true", help="Run sync on any stable .xlsm change, not just email diffs")
    parser.add_argument("--template", default=None)
    parser.add_argument("--no-macro", default=None)
    args = parser.parse_args()

    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)
    template = args.template or os.path.join(project_root, "Pending Pickup Template.xlsm")
    no_macro = args.no_macro or os.path.join(project_root, "Pending Pickup Template (no-macros).xlsx")
    sync_script = os.path.join(script_dir, "sync_template.py")

    log_dir = os.path.join(project_root, "log")
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir, "template-watch.log")

    if not os.path.exists(template):
        log(f"ERROR: template not found: {template}", log_path)
        return 2

    last_mtime = file_stat(template)
    last_sig = None
    sig, err = read_email_config(template)
    if err:
        log(f"WARN: initial read failed: {err}", log_path)
    else:
        last_sig = sig
        log("Initial email config read from .xlsm OK.", log_path)

    log(f"Watching {template} (interval={args.interval}s, settle={args.settle}s). Ctrl+C to stop.", log_path)

    try:
        while True:
            time.sleep(args.interval)
            cur = file_stat(template)
            if cur is None or cur == last_mtime:
                continue

            # changed; wait for stable save
            if not is_stable(template, args.settle):
                log("Detected change but file not stable yet; will retry.", log_path)
                continue

            last_mtime = file_stat(template)
            log("Detected stable change in .xlsm; checking email sections...", log_path)

            xlsm_sig, err1 = read_email_config(template)
            if err1:
                log(f"Could not read .xlsm (may be locked by Excel): {err1}", log_path)
                continue

            need_sync = bool(args.always_sync_on_change)
            reason = "forced (--always-sync-on-change)" if need_sync else None

            if not need_sync:
                xlsx_sig, err2 = read_email_config(no_macro)
                if err2:
                    need_sync = True
                    reason = f"no-macro unreadable/missing: {err2}"
                elif xlsm_sig != xlsx_sig:
                    need_sync = True
                    # summarize diffs
                    diffs = []
                    for sh in EMAIL_SHEETS:
                        for c in EMAIL_CELLS:
                            a = (xlsm_sig.get(sh) or {}).get(c)
                            b = (xlsx_sig.get(sh) or {}).get(c)
                            if a != b:
                                diffs.append(f"{sh}!{c}")
                    reason = "email config differs: " + ", ".join(diffs[:8]) + ("..." if len(diffs) > 8 else "")
                else:
                    reason = "email config already matches no-macro"

            if not need_sync:
                log(f"No sync needed ({reason}).", log_path)
                last_sig = xlsm_sig
                continue

            log(f"Running sync_template.py ({reason})...", log_path)
            try:
                r = subprocess.run([sys.executable, sync_script], cwd=project_root, capture_output=True, text=True, timeout=30)
                if r.returncode == 0:
                    log("sync_template.py completed successfully.", log_path)
                    if r.stdout.strip():
                        for ln in r.stdout.strip().splitlines()[-12:]:
                            log("  " + ln, log_path)
                else:
                    log(f"sync_template.py failed (exit {r.returncode}).", log_path)
                    out = (r.stdout or "") + ("\n" + r.stderr if r.stderr else "")
                    for ln in out.strip().splitlines()[-20:]:
                        log("  " + ln, log_path)
            except subprocess.TimeoutExpired:
                log("sync_template.py timed out (>30s). Excel/process may need cleanup; investigate logs before retrying.", log_path)
                try:
                    subprocess.run(["taskkill", "/F", "/IM", "excel.exe"], capture_output=True, timeout=10)
                except Exception:
                    pass
            except Exception as e:
                log(f"Failed to run sync_template.py: {e}", log_path)

            last_sig = xlsm_sig

    except KeyboardInterrupt:
        log("Watcher stopped.", log_path)
        return 0


if __name__ == "__main__":
    sys.exit(main())
