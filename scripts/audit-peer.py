# audit-peer.py
# Compares the shared, portable parts of this project against a sibling project and
# reports what each side has that the other does not: files, VBA procedures,
# Python function names, and Markdown sections.
#
# It never changes either project. It only reads.
#
#   python scripts/audit-peer.py
#   python scripts/audit-peer.py --sections scripts,commands
#   python scripts/audit-peer.py --peer "C:\path\to\Other Project" --report documentation\cross-audit.md
#
# The peer is found automatically: a sibling directory that holds a root *.bas and a
# scripts\ folder. Pass --peer when the projects are not siblings.
#
# Exit codes: 0 = audited, 2 = peer not found, 3 = nothing auditable on either side.

import argparse
import hashlib
import os
import re
import sys
from datetime import datetime

project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

section_specs = {
    "scripts":       {"dir": "scripts",          "include": ["*.py", "*.cmd", "*.bat"], "recursive": False},
    "documentation": {"dir": "documentation",    "include": ["*.md", "*.py"],           "recursive": False},
    "commands":      {"dir": ".kilo/command",    "include": ["*.md"],                            "recursive": False},
    "vba":           {"dir": ".",                "include": ["*.bas"],                           "recursive": False},
    "rules":         {"dir": "documentation",    "include": ["AGENTS.md"],                       "recursive": False},
}


def resolve_peer_root(requested):
    if requested:
        if not os.path.exists(requested):
            print(f"Peer not found: {requested}", file=sys.stderr)
            sys.exit(2)
        return os.path.abspath(requested)
    parent = os.path.dirname(project_root)
    candidates = []
    if os.path.isdir(parent):
        for entry in os.listdir(parent):
            full = os.path.join(parent, entry)
            if not os.path.isdir(full) or full == project_root:
                continue
            has_bas = any(f.endswith(".bas") for f in os.listdir(full) if os.path.isfile(os.path.join(full, f)))
            has_scripts = os.path.isdir(os.path.join(full, "scripts"))
            if has_bas and has_scripts:
                candidates.append(full)
    if len(candidates) == 1:
        return candidates[0]
    if len(candidates) == 0:
        print(
            "No sibling project with a root *.bas and a scripts\\ folder was found next to "
            f"{project_root}",
            file=sys.stderr,
        )
        print("Pass --peer <path to the other project>.", file=sys.stderr)
        sys.exit(2)
    print("Several candidate peers found; pick one with --peer:", file=sys.stderr)
    for c in candidates:
        print(f"  {c}", file=sys.stderr)
    sys.exit(2)


def get_section_files(root, wanted):
    files = []
    for section in wanted:
        spec = section_specs.get(section)
        if not spec:
            print(
                f"Unknown section '{section}' (known: {', '.join(section_specs.keys())})",
                file=sys.stderr,
            )
            sys.exit(2)
        dir_path = os.path.join(root, spec["dir"])
        if not os.path.isdir(dir_path):
            continue
        for f in os.listdir(dir_path):
            fpath = os.path.join(dir_path, f)
            if not os.path.isfile(fpath):
                continue
            matched = False
            for pattern in spec["include"]:
                import fnmatch
                if fnmatch.fnmatch(f, pattern):
                    matched = True
                    break
            if not matched:
                continue
            rel = os.path.relpath(fpath, root)
            kind = os.path.splitext(f)[1].lstrip(".").lower()
            files.append({
                "section": section,
                "rel": rel,
                "name": f,
                "kind": kind,
                "path": fpath,
                "size": os.path.getsize(fpath),
                "hash": hashlib.sha256(open(fpath, "rb").read()).hexdigest(),
                "lines": 0,
                "symbols": [],
            })
    return files


symbol_rx = {
    "ps1": r'^\s*function\s+([A-Za-z_][\w-]*)',
    "py": r'^\s*(?:def|class)\s+(\w+)',
    "bas": r'^\s*(?:(?:Public|Private|Friend|Static)\s+)*(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)',
    "md": r'^#{1,4}\s+(.+?)\s*#*\s*$',
}


def get_symbol_list(path, kind):
    rx = symbol_rx.get(kind)
    if not rx:
        return []
    found = []
    seen = set()
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        for line in f:
            m = re.match(rx, line)
            if m:
                name = m.group(1).strip()
                if name and name.lower() not in seen:
                    seen.add(name.lower())
                    found.append(name)
    return found


def get_match_for(file_obj, pool):
    for other in pool:
        if other["rel"] == file_obj["rel"]:
            return other
    for other in pool:
        if other["name"] == file_obj["name"]:
            return other
    return None


def format_list(items, limit, prefix="      "):
    if not items:
        return []
    shown = []
    for i in range(min(len(items), limit)):
        shown.append(f"{prefix}{items[i]}")
    if len(items) > limit:
        shown.append(f"{prefix}... (+{len(items) - limit} more)")
    return shown


def main():
    parser = argparse.ArgumentParser(description="Cross-project audit of portable files")
    parser.add_argument("--peer", default=None)
    parser.add_argument("--sections", nargs="*", default=["scripts", "documentation", "commands", "vba", "rules"])
    parser.add_argument("--report", default=None)
    parser.add_argument("--symbol-limit", type=int, default=20)
    parser.add_argument("--quiet", action="store_true")
    args = parser.parse_args()

    here_files = get_section_files(project_root, args.sections)
    peer_root = resolve_peer_root(args.peer)
    peer_leaf = os.path.basename(peer_root)
    this_leaf = os.path.basename(project_root)
    peer_files = get_section_files(peer_root, args.sections)

    if not here_files and not peer_files:
        print(
            f"Nothing auditable in either project for sections: {', '.join(args.sections)}",
            file=sys.stderr,
        )
        sys.exit(3)

    for group in (here_files, peer_files):
        for f in group:
            with open(f["path"], "r", encoding="utf-8", errors="replace") as fh:
                f["lines"] = sum(1 for _ in fh)
            f["symbols"] = get_symbol_list(f["path"], f["kind"])

    rows = []
    port_candidates = []
    port_back = []

    for peer in peer_files:
        match = get_match_for(peer, here_files)
        added_peer = []
        added_here = []
        if not match:
            status = "peer only"
            detail = f"{peer['lines']} lines"
            port_candidates.append(peer)
        else:
            added_here = []
            added_peer = []
            for s in peer["symbols"]:
                if s not in match["symbols"]:
                    added_peer.append(s)
            for s in match["symbols"]:
                if s not in peer["symbols"]:
                    added_here.append(s)
            if peer["hash"] == match["hash"]:
                status = "identical"
                detail = f"{peer['lines']} lines"
            else:
                status = "changed"
                detail = f"peer {peer['lines']} vs here {match['lines']} lines"
                if added_peer:
                    detail += " | peer adds: " + ", ".join(added_peer[:3])
                if added_here:
                    detail += " | here adds: " + ", ".join(added_here[:3])
        rows.append({
            "section": peer["section"],
            "file": peer["rel"],
            "status": status,
            "detail": detail,
            "added": added_peer,
            "match": match,
        })

    for here in here_files:
        match = get_match_for(here, peer_files)
        if not match:
            port_back.append(here)
            rows.append({
                "section": here["section"],
                "file": here["rel"],
                "status": "here only",
                "detail": f"{here['lines']} lines",
                "added": [],
                "match": None,
            })

    port_back_rows = [r for r in rows if r["status"] == "here only"]

    md = []
    md.append(f"# Cross audit: {this_leaf} <-> {peer_leaf}")
    md.append("")
    md.append(f"Generated {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    md.append("")
    md.append(f"- this: `{project_root}`")
    md.append(f"- peer: `{peer_root}`")

    for section in args.sections:
        section_rows = [r for r in rows if r["section"] == section]
        if not section_rows:
            continue
        if not args.quiet:
            print("")
            print(f"== {section.upper()} ==")
        md.append(f"## {section}")
        md.append("")
        md.append("| file | status | detail |")
        md.append("| --- | --- | --- |")
        for r in section_rows:
            colour = "Gray"
            if r["status"] == "peer only":
                colour = "Green"
            if r["status"] == "here only":
                colour = "Magenta"
            if r["status"] == "changed":
                colour = "Yellow"
            if not args.quiet:
                print(f"  {r['file']:<46} {r['status']:<11} {r['detail']}")
                if len(r["added"]) > 3:
                    for line in format_list(r["added"], args.symbol_limit):
                        if line.strip():
                            print(line)
            md.append(f"| `{r['file']}` | {r['status']} | {r['detail']} |")
            for a in r["added"][:args.symbol_limit]:
                md.append(f"| | | + {a} |")

    # -------------------------------------------------- VBA procedure cross-diff
    vba_here = {}
    for f in [x for x in here_files if x["section"] == "vba"]:
        for s in f["symbols"]:
            if s.lower() not in vba_here:
                vba_here[s.lower()] = s
    vba_peer = {}
    for f in [x for x in peer_files if x["section"] == "vba"]:
        for s in f["symbols"]:
            if s.lower() not in vba_peer:
                vba_peer[s.lower()] = s
    vba_peer_only = []
    vba_shared = []
    if vba_peer or vba_here:
        for key in vba_peer:
            if key in vba_here:
                vba_shared.append(vba_peer[key])
            else:
                vba_peer_only.append(vba_peer[key])
        vba_peer_only = sorted(vba_peer_only)
        vba_shared = sorted(vba_shared)
        if not args.quiet:
            print("")
            print(f"== VBA PROCEDURES ==  shared={len(vba_shared)}  peer-only={len(vba_peer_only)}  here={len(vba_here)} ==")
            for line in format_list(vba_peer_only, args.symbol_limit, "  peer only  "):
                print(line)
            for line in format_list(vba_shared, 10, "  shared     "):
                print(line)
        md.append("## VBA procedures")
        md.append("")
        md.append(f"- shared: {len(vba_shared)}")
        md.append(f"- peer only: {len(vba_peer_only)}")
        md.append(f"- this project: {len(vba_here)}")
        md.append("")
        md.append("### Procedures only the peer has (port candidates)")
        md.append("")
        for s in vba_peer_only:
            md.append(f"- {s}")

    # ------------------------------------------------------------- port shortlists
    if port_candidates:
        script_ports = [f for f in port_candidates if f["section"] in ("scripts", "documentation")]
        if script_ports:
            if not args.quiet:
                print("")
                print("== WORTH A LOOK (peer has, this project does not) ==")
            md.append("## Port candidates (peer -> this)")
            md.append("")
            for f in script_ports:
                line = f"  {f['rel']:<46} {f['lines'],6} lines"
                if not args.quiet:
                    print(line)
                md.append(f"- `{f['rel']}` - {f['lines']} lines, {round(f['size'] / 1024, 1)} KB")
                for s in f["symbols"][:args.symbol_limit]:
                    if not args.quiet:
                        print(f"      {s}")
                    md.append(f"  - {s}")

        cmd_ports = [f for f in port_candidates if f["section"] == "commands"]
        if cmd_ports:
            if not args.quiet:
                print("")
                print(f"Commands only the peer has: {', '.join(f['name'] for f in cmd_ports)}")
            md.append("### Commands only the peer has")
            md.append("")
            for f in cmd_ports:
                md.append(f"- `{f['rel']}`")

        vba_ports = [f for f in port_candidates if f["section"] == "vba"]
        if vba_ports:
            if not args.quiet:
                print(f"VBA modules only the peer has: {', '.join(f['name'] for f in vba_ports)}")
            md.append("### VBA modules only the peer has")
            md.append("")
            for f in vba_ports:
                md.append(f"- `{f['rel']}` ({f['lines']} lines, {len(f['symbols'])} procedures)")

    if port_back_rows:
        if not args.quiet:
            print("")
            print("== PEER MAY WANT (this has, peer does not) ==")
            for r in port_back_rows:
                print(f"  {r['file']:<46} {r['detail']}")
        md.append("## Port candidates (this -> peer)")
        md.append("")
        for r in port_back_rows:
            md.append(f"- `{r['file']}` - {r['detail']}")

    # --------------------------------------------------------------------- summary
    identical = sum(1 for r in rows if r["status"] == "identical")
    changed = sum(1 for r in rows if r["status"] == "changed")
    peer_only = sum(1 for r in rows if r["status"] == "peer only")
    here_only = sum(1 for r in rows if r["status"] == "here only")

    if not args.quiet:
        print("")
        print(f"SUMMARY  identical={identical}  changed={changed}  peer-only={peer_only}  here-only={here_only}")

    md.append("")
    md.append("## Summary")
    md.append("")
    md.append(f"- identical: {identical}")
    md.append(f"- changed: {changed}")
    md.append(f"- peer only: {peer_only}")
    md.append(f"- here only: {here_only}")
    md.append("")
    md.append("This is a read-only report. Porting is a manual decision: copy a file, adapt its")
    md.append("data source and paths, then run this project's compile gate and report build.")
    md.append("")

    if args.report:
        report_path = args.report
        if not os.path.isabs(report_path):
            report_path = os.path.join(project_root, report_path)
        rdir = os.path.dirname(report_path)
        if rdir and not os.path.isdir(rdir):
            os.makedirs(rdir, exist_ok=True)
        with open(report_path, "w", encoding="utf-8") as f:
            f.write("\n".join(md))
        if not args.quiet:
            print("")
            print(f"Report written: {report_path}")

    sys.exit(0)


if __name__ == "__main__":
    main()
