import difflib
import os

root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
src_dir = os.path.join(root, "vba")
tpl_dir = os.path.join(root, "scripts", "temp", "tpl_vba")


def norm(path):
    raw = open(path, encoding="utf-8", errors="replace").read().replace("\r\n", "\n")
    return [ln.rstrip() for ln in raw.split("\n") if ln.strip()]


for f in sorted(os.listdir(src_dir)):
    if not f.endswith(".bas"):
        continue
    a = norm(os.path.join(src_dir, f))
    b = norm(os.path.join(tpl_dir, f + ".txt"))
    diff = [d for d in difflib.unified_diff(a, b, fromfile="src/" + f, tofile="tpl/" + f, lineterm="", n=2)]
    if diff:
        print(f"\n=== {f} ({len(diff)} diff lines)")
        print("\n".join(diff))
    else:
        print(f"{f}: identical (ignoring blank lines)")
