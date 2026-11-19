import os
import re

d = r"C:\Users\smd32\Downloads\Pending Pickup Project\vba"
proc_re = re.compile(r"^(?:(?:Public|Private|Friend)\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\b", re.I)
decl_re = re.compile(r"^(?:(?:Public|Private|Global)\s+)+(?:Const\s+)?([A-Za-z_]\w*)", re.I)

for f in sorted(os.listdir(d)):
    if not f.endswith(".bas"):
        continue
    lines = open(os.path.join(d, f), encoding="utf-8").read().split("\n")
    first_proc = None
    late = []
    for i, ln in enumerate(lines):
        if proc_re.match(ln.strip()):
            if first_proc is None:
                first_proc = i
            continue
        if first_proc is not None and decl_re.match(ln.strip()) and not ln.startswith(("'", " ")):
            late.append((i + 1, ln.strip()))
    print(f"{f}: first_proc_line={first_proc + 1 if first_proc is not None else None} late_decls={len(late)}")
    for ln, txt in late[:20]:
        print(f"   {ln}: {txt}")
