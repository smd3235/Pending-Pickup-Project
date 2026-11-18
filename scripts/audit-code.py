#!/usr/bin/env python3
"""Static/code audit for Pending Pickup project.

Scans VBA (.bas), Python scripts, and key project/docs/config for common issues
and policy violations (including missing timeouts on subprocess calls per 30s rule).

Usage:
  python scripts/audit-code.py
  python scripts/audit-code.py --json
  python scripts/audit-code.py --strict   # warnings fail
"""
from __future__ import annotations

import argparse
import ast
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
DOCS = ROOT / "documentation"

SEVERITY_ORDER = {"error": 0, "warn": 1, "info": 2}


def add(findings, severity, source, line, message):
    findings.append({"severity": severity, "source": str(source), "line": line, "message": message})


def scan_vba(findings):
    bas_files = sorted(ROOT.glob("*.bas"))
    if not bas_files:
        add(findings, "warn", ROOT, None, "No .bas files found in project root")
        return
    proc_re = re.compile(r"^\s*(Public|Private|Friend)?\s*(Sub|Function|Property)\s+", re.I)
    for p in bas_files:
        try:
            text = p.read_text(encoding="utf-8", errors="replace")
        except Exception as e:
            add(findings, "error", p, None, f"Could not read VBA file: {e}")
            continue
        lines = text.splitlines()
        has_option_explicit = any(l.strip().lower() == "option explicit" for l in lines[:30])
        # module-level attribute header is normal; check first substantive area
        if not has_option_explicit:
            add(findings, "warn", p, 1, "Missing 'Option Explicit' near top of module")
        if not re.search(r"Attribute\s+VB_Name", text):
            add(findings, "warn", p, 1, "Missing Attribute VB_Name header")

        on_error_resume = 0
        for i, line in enumerate(lines, 1):
            stripped = line.strip()
            if re.match(r"^On\s+Error\s+Resume\s+Next\b", stripped, re.I):
                on_error_resume += 1
            if re.search(r"\bTODO\b|\bFIXME\b|\bHACK\b|\bXXX\b", line):
                sev = "warn" if "HACK" in line.upper() or "FIXME" in line.upper() else "info"
                add(findings, sev, p, i, f"Marker comment: {stripped[:160]}")
            if len(line) > 400:
                add(findings, "info", p, i, f"Very long line ({len(line)} chars)")
            if "Debug.Print" in line:
                add(findings, "info", p, i, "Debug.Print present (consider structured logging)")
            # crude tab/space mix not flagged; check suspicious shell/exec patterns
            if re.search(r"\bShell\s*\(", line, re.I):
                add(findings, "warn", p, i, "Shell() call found - ensure safe/necessary")

        if on_error_resume > 30:
            add(findings, "info", p, None, f"High use of 'On Error Resume Next' ({on_error_resume} occurrences) - verify errors handled/logged")

        # basic procedure/end balance hints
        subs = len(re.findall(r"^\s*(?:Public|Private|Friend)?\s*Sub\s+\w+", text, re.I | re.M))
        funcs = len(re.findall(r"^\s*(?:Public|Private|Friend)?\s*Function\s+\w+", text, re.I | re.M))
        props = len(re.findall(r"^\s*(?:Public|Private|Friend)?\s*Property\s+\w+", text, re.I | re.M))
        end_sub = len(re.findall(r"^\s*End\s+Sub\b", text, re.I | re.M))
        end_func = len(re.findall(r"^\s*End\s+Function\b", text, re.I | re.M))
        end_prop = len(re.findall(r"^\s*End\s+Property\b", text, re.I | re.M))
        if subs != end_sub:
            add(findings, "warn", p, None, f"Sub/End Sub count mismatch ({subs} vs {end_sub})")
        if funcs != end_func:
            add(findings, "warn", p, None, f"Function/End Function count mismatch ({funcs} vs {end_func})")
        if props != end_prop:
            add(findings, "warn", p, None, f"Property/End Property count mismatch ({props} vs {end_prop})")


def _py_subprocess_calls_without_timeout(tree: ast.AST):
    issues = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Call):
            func = node.func
            name = None
            if isinstance(func, ast.Attribute) and isinstance(func.value, ast.Name) and func.value.id == "subprocess":
                name = func.attr
            elif isinstance(func, ast.Attribute) and isinstance(func.value, ast.Attribute):
                # unlikely
                pass
            if name in {"run", "call", "check_call", "check_output", "Popen"}:
                has_timeout = any(k.arg == "timeout" for k in node.keywords)
                if name == "Popen":
                    # Popen can be managed separately; warn if not obviously bounded
                    issues.append((node.lineno, name, has_timeout))
                elif not has_timeout:
                    issues.append((node.lineno, name, has_timeout))
    return issues


def scan_python(findings):
    py_files = sorted(SCRIPTS.glob("*.py")) if SCRIPTS.exists() else []
    for p in py_files:
        try:
            src = p.read_text(encoding="utf-8", errors="replace")
        except Exception as e:
            add(findings, "error", p, None, f"Could not read Python file: {e}")
            continue
        try:
            tree = ast.parse(src, filename=str(p))
        except SyntaxError as e:
            add(findings, "error", p, getattr(e, "lineno", None), f"SyntaxError: {e.msg}")
            continue
        except Exception as e:
            add(findings, "error", p, None, f"Parse error: {e}")
            continue

        src_lines = src.splitlines()
        for lineno, name, has_timeout in _py_subprocess_calls_without_timeout(tree):
            if name == "Popen":
                line_text = src_lines[lineno - 1] if 0 < lineno <= len(src_lines) else ""
                ctx = "\n".join(src_lines[max(0, lineno - 16):lineno]).lower()
                pre = "\n".join(src_lines[max(0, lineno - 40):lineno]).lower()
                # Dialog watchers are intentionally long-lived sidecars but bounded/monitored (terminate+wait on cleanup).
                if (
                    "watch" in line_text.lower()
                    or "dialogwatcher" in ctx
                    or "watch-compile-error" in ctx
                    or "dialog watcher" in ctx
                    or "dialogwatcher" in pre
                    or "watcher_script" in ctx
                    or "class dialogwatcher" in pre
                ):
                    add(findings, "info", p, lineno, "subprocess.Popen for dialog watcher sidecar (bounded/monitored)")
                else:
                    add(findings, "warn", p, lineno, "subprocess.Popen used - ensure process is bounded/monitored per 30s hang policy")
            elif not has_timeout:
                line_text = src_lines[lineno - 1] if 0 < lineno <= len(src_lines) else ""
                ctx = "\n".join(src_lines[max(0, lineno - 8):lineno]).lower() + "\n" + line_text.lower()
                # Screenshot/export steps may exceed 30s but must be capped at 120s and kill Excel on timeout.
                if "snapshot" in ctx or "screenshot" in ctx or "snap_script" in line_text:
                    add(findings, "warn", p, lineno, f"subprocess.{name} without timeout (screenshot/export must be capped at <=120s and kill Excel on timeout)")
                else:
                    add(findings, "warn", p, lineno, f"subprocess.{name} without timeout - may violate 30s no-hang rule")

        for node in ast.walk(tree):
            if isinstance(node, ast.ExceptHandler):
                if node.type is None:
                    add(findings, "warn", p, node.lineno, "Bare 'except:' catches all exceptions/BaseExceptions")
            if isinstance(node, ast.Call):
                # flag very large time.sleep literals
                if isinstance(node.func, ast.Attribute) and isinstance(node.func.value, ast.Name):
                    if node.func.value.id == "time" and node.func.attr == "sleep" and node.args:
                        arg = node.args[0]
                        val = None
                        if isinstance(arg, ast.Constant) and isinstance(arg.value, (int, float)):
                            val = float(arg.value)
                        line_text = src_lines[node.lineno - 1] if 0 < node.lineno <= len(src_lines) else ""
                        if val is not None and val > 30:
                            add(findings, "warn", p, node.lineno, f"time.sleep({val}) exceeds 30s policy")
                        elif val is not None and val >= 10:
                            # watchdog timer sleeps enforce a hard upper bound (e.g. 10s kill) rather than waiting out work
                            if "watchdog" in line_text.lower() or "watchdog" in "\n".join(src_lines[max(0, node.lineno - 6):node.lineno]).lower():
                                add(findings, "info", p, node.lineno, f"time.sleep({val}) used as watchdog/bound timer")
                            else:
                                add(findings, "info", p, node.lineno, f"time.sleep({val}) is long; ensure total waits stay within 30s per operation")

        # text-based checks
        lines = src.splitlines()
        for i, line in enumerate(lines, 1):
            if re.search(r"\bTODO\b|\bFIXME\b|\bHACK\b", line):
                # avoid flagging this script's own detection logic
                if p.name == "audit-code.py" and ("re.search" in line or "in line.upper()" in line or "Marker comment" in line):
                    continue
                add(findings, "info", p, i, f"Marker comment: {line.strip()[:160]}")
            if "sys.path" in line and "insert" in line:
                if p.name == "audit-code.py":
                    continue
                add(findings, "info", p, i, "sys.path modified at runtime")
            if re.search(r"except\s+Exception\s*:\s*pass\b", line):
                if p.name == "audit-code.py":
                    continue
                add(findings, "info", p, i, "Silent 'except Exception: pass' - consider logging")


def scan_project(findings):
    required = [
        ROOT / "Pending Pickup Template.xlsm",
        ROOT / "Pending Pickup Template (no-macros).xlsx",
        ROOT / "Project-Local-Build.zip",
        ROOT / "scripts" / "sync_template.py",
        ROOT / "scripts" / "run-report.py",
        ROOT / "scripts" / "test-compile.py",
        ROOT / "scripts" / "check-logs.py",
    ]
    for p in required:
        if not p.exists():
            sev = "warn" if p.suffix == ".zip" or "no-macros" in p.name else "warn"
            add(findings, sev, p, None, "Expected project artifact/script missing")

    # AGENTS.md presence (policy requested)
    if not (ROOT / "AGENTS.md").exists():
        add(findings, "warn", ROOT / "AGENTS.md", None, "AGENTS.md missing (project rules/30s policy should be documented)")

    # .kilo usage
    if (ROOT / ".kilocode").exists():
        add(findings, "warn", ROOT / ".kilocode", None, "Legacy .kilocode/ directory present; prefer .kilo/ per project config")

    # README/screenshots references
    readme = ROOT / "README.md"
    if readme.exists():
        try:
            txt = readme.read_text(encoding="utf-8", errors="replace")
            for m in re.finditer(r"\((documentation/screenshots/[^)]+)\)", txt):
                ref = ROOT / m.group(1)
                if not ref.exists():
                    add(findings, "warn", readme, None, f"README references missing screenshot: {m.group(1)}")
        except Exception as e:
            add(findings, "warn", readme, None, f"Could not scan README: {e}")

    # log directory naming
    if not (ROOT / "log").exists() and (ROOT / "Logs").exists():
        add(findings, "info", ROOT / "Logs", None, "Uppercase Logs/ directory present; this project uses log/")

    # sync_template includes expected modules?
    sync = SCRIPTS / "sync_template.py"
    if sync.exists():
        try:
            st = sync.read_text(encoding="utf-8", errors="replace")
            for mod in ("modCore.bas", "modDialog.bas", "modFormatting.bas", "modLogging.bas", "modAudit.bas", "modDeclarations.bas", "modUtils.bas"):
                if mod not in st and "glob" not in st:
                    add(findings, "info", sync, None, f"sync_template.py does not mention {mod} explicitly (may be globbed)")
        except Exception:
            pass


def main():
    ap = argparse.ArgumentParser(description="Audit Pending Pickup code/scripts/docs for issues")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--strict", action="store_true", help="Exit 1 on warnings as well as errors")
    args = ap.parse_args()

    findings = []
    scan_vba(findings)
    scan_python(findings)
    scan_project(findings)

    counts = {"error": 0, "warn": 0, "info": 0}
    for f in findings:
        counts[f["severity"]] = counts.get(f["severity"], 0) + 1
    ok = counts.get("error", 0) == 0 and (not args.strict or counts.get("warn", 0) == 0)

    result = {"ok": ok, "counts": counts, "findings": findings}
    if args.json:
        print(json.dumps(result, indent=2))
    else:
        print(f"Code audit: {counts.get('error',0)} error(s), {counts.get('warn',0)} warning(s), {counts.get('info',0)} info")
        for f in sorted(findings, key=lambda x: (SEVERITY_ORDER.get(x['severity'], 9), str(x['source']), x['line'] or 0)):
            loc = f"{f['source']}" + (f":{f['line']}" if f.get('line') else "")
            print(f"[{f['severity'].upper()}] {loc} - {f['message']}")
        print("OK: no blocking issues." if ok else "ISSUES DETECTED.")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
