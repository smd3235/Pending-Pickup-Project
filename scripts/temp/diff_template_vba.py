import os

from oletools.olevba import VBA_Parser

root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
tpl = os.path.join(root, "Pending Pickup Template.xlsm")
out_dir = os.path.join(root, "scripts", "temp", "tpl_vba")
os.makedirs(out_dir, exist_ok=True)

p = VBA_Parser(tpl)
for _fn, _sp, name, code in p.extract_macros():
    path = os.path.join(out_dir, name + ".txt")
    with open(path, "w", encoding="utf-8") as f:
        f.write(code)
    print(name, len(code))

src_dir = os.path.join(root, "vba")
import difflib

for f in sorted(os.listdir(src_dir)):
    if not f.endswith(".bas"):
        continue
    tpl_path = os.path.join(out_dir, f + ".txt")
    if not os.path.exists(tpl_path):
        print(f"{f}: MISSING")
        continue
    a = open(os.path.join(src_dir, f), encoding="utf-8", errors="replace").read().replace("\r\n", "\n").splitlines()
    b = open(tpl_path, encoding="utf-8", errors="replace").read().replace("\r\n", "\n").splitlines()
    diff = list(difflib.unified_diff(a, b, fromfile="src/" + f, tofile="tpl/" + f, lineterm="", n=1))
    if diff:
        print(f"\n=== {f} DIFF ({len(diff)} lines)")
        print("\n".join(diff[:80]))
    else:
        print(f"{f}: identical")
