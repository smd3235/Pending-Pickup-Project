import os

import psutil
import pythoncom
import win32com.client

root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
src_dir = os.path.join(root, "vba")
work = os.path.join(root, "scripts", "temp", "decl_check.xlsm")
temp_bas = os.path.join(root, "scripts", "temp", "modMin2.bas")

for p in psutil.process_iter(["name"]):
    if p.info["name"] and p.info["name"].lower() == "excel.exe":
        p.kill()

with open(temp_bas, "w", encoding="utf-8", newline="\n") as f:
    f.write('Attribute VB_Name = "modMin"\nOption Explicit\n\nPublic Sub Probe1()\n    Debug.Print mInReportRun\nEnd Sub\n')

pythoncom.CoInitialize()
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
try:
    wb = excel.Workbooks.Add()
    wb.SaveAs(work, FileFormat=52)
    vbp = wb.VBProject
    vbp.VBComponents.Import(os.path.join(src_dir, "modDeclarations.bas"))
    vbp.VBComponents.Import(temp_bas)
    cm = vbp.VBComponents("modDeclarations").CodeModule
    print("declarations lines:", cm.CountOfLines)
    print("---- tail ----")
    print(cm.Lines(max(1, cm.CountOfLines - 30), 30))
    print("---- names of public vars present ----")
    txt = cm.Lines(1, cm.CountOfLines)
    for nm in ["mInReportRun", "mRunLogPath", "mAuditIssues", "mReportStartTime"]:
        print(nm, "present:", nm in txt)
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
