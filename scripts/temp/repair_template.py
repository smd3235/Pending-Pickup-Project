import os
import shutil
import time

import psutil
import pythoncom
import win32com.client

root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
template = os.path.join(root, "Pending Pickup Template.xlsm")
backup = os.path.join(root, "scripts", "temp", "template_before_repair.xlsm")
src_dir = os.path.join(root, "vba")

for p in psutil.process_iter(["name"]):
    if p.info["name"] and p.info["name"].lower() == "excel.exe":
        p.kill()
time.sleep(1)

shutil.copy2(template, backup)
print("backup:", backup, flush=True)

pythoncom.CoInitialize()
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
try:
    wb = excel.Workbooks.Open(template)
    vbp = wb.VBProject
    for f in sorted(os.listdir(src_dir)):
        if not f.endswith(".bas"):
            continue
        name = f[:-4]
        text = open(os.path.join(src_dir, f), encoding="utf-8").read().replace("\r\n", "\n")
        lines = [ln for ln in text.split("\n") if not ln.startswith("Attribute VB_Name")]
        code = "\n".join(lines).rstrip() + "\n"
        cm = vbp.VBComponents(name).CodeModule
        cm.DeleteLines(1, cm.CountOfLines)
        cm.AddFromString(code)
        print(f"rewrote {name}: {cm.CountOfLines} lines", flush=True)
    wb.Save()
    wb.Close(SaveChanges=False)
    print("template saved", flush=True)
finally:
    try:
        excel.Quit()
    except Exception:
        pass
    pythoncom.CoUninitialize()
    for p in psutil.process_iter(["name"]):
        if p.info["name"] and p.info["name"].lower() == "excel.exe":
            p.kill()
