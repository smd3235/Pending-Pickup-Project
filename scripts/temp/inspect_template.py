import os
import re
import zipfile

from oletools.olevba import VBA_Parser

root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
tpl = os.path.join(root, "Pending Pickup Template.xlsm")

z = zipfile.ZipFile(tpl)
for n in z.namelist():
    if n.startswith("xl/drawings/") and n.endswith(".xml"):
        d = z.read(n).decode("utf-8", "replace")
        names = re.findall(r'<xdr:cNvPr[^>]*name="([^"]+)"[^>]*/?>', d)
        macros = re.findall(r'macro="([^"]*)"', d)
        print(f"--- {n}")
        for nm in names:
            print("  shape:", nm)
        for m in macros:
            print("  macro:", m)

print("\n=== VBA modules in template ===")
p = VBA_Parser(tpl)
mods = {}
for _fn, _sp, vba_name, code in p.extract_macros():
    mods[vba_name] = code
    print(f"  {vba_name:24s} {len(code):7d} bytes")

print("\n=== Compare template modules to vba/*.bas ===")
vba_dir = os.path.join(root, "vba")
for f in sorted(os.listdir(vba_dir)):
    if not f.endswith(".bas"):
        continue
    src = open(os.path.join(vba_dir, f), encoding="utf-8", errors="replace").read()
    key = f
    tpl_code = mods.get(key)
    if tpl_code is None:
        print(f"  {f:26s} MISSING from template")
    elif tpl_code.replace("\r\n", "\n").strip() == src.replace("\r\n", "\n").strip():
        print(f"  {f:26s} identical")
    else:
        print(f"  {f:26s} DIFFERS (src={len(src)}, tpl={len(tpl_code)})")

mono = os.path.join(root, "Pending_Pickup_Macro.bas")
print("monolithic module present in template:", "PendingPickupReport" in mods or "Pending_Pickup_Macro" in mods)
