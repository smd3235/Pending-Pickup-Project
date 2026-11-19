import difflib
import os
import shutil

import psutil
import pythoncom
import win32com.client

root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
template = os.path.join(root, "Pending Pickup Template.xlsm")
test_copy = os.path.join(root, "scripts", "temp", "code_dump.xlsm")
src_dir = os.path.join(root, "vba")

for p in psutil.process_iter(["name"]):
    if p.info["name"] and p.info["name"].lower() == "excel.exe":
        p.kill()

if os.path.exists(test_copy):
    os.remove(test_copy)
shutil.copy2(template, test_copy)

pythoncom.CoInitialize()
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
live = {}
try:
    wb = excel.Workbooks.Open(test_copy)
    for comp in wb.VBProject.VBComponents:
        if comp.Type == 1:
            live[comp.Name] = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
    wb.Close(SaveChanges=False)
finally:
    try:
        excel.Quit()
    except Exception:
        pass
    pythoncom.CoUninitialize()
    for p in psutil.process_iter(["name"]):
        if p.info["name"] and p.info["name"].lower() == "excel.exe":
            p.kill()

out_dir = os.path.join(root, "scripts", "temp", "live_vba")
os.makedirs(out_dir, exist_ok=True)
for name, code in live.items():
    with open(os.path.join(out_dir, name + ".bas"), "w", encoding="utf-8", newline="\n") as f:
        f.write(code.replace("\r\n", "\n"))

for f in sorted(os.listdir(src_dir)):
    if not f.endswith(".bas"):
        continue
    src = open(os.path.join(src_dir, f), encoding="utf-8", errors="replace").read().replace("\r\n", "\n")
    live_path = os.path.join(out_dir, f)
    if not os.path.exists(live_path):
        print(f"{f}: MISSING IN TEMPLATE")
        continue
    tpl = open(live_path, encoding="utf-8", errors="replace").read()
    if src == tpl:
        print(f"{f}: identical")
        continue
    diff = list(difflib.unified_diff(src.splitlines(), tpl.splitlines(), fromfile="src/" + f, tofile="tpl/" + f, lineterm="", n=1))
    print(f"\n=== {f}: {len(diff)} diff lines")
    print("\n".join(diff[:120]))
