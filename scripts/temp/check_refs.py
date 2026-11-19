import ctypes
import ctypes.wintypes as wt
import os
import shutil

import psutil
import pythoncom
import win32com.client

root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
template = os.path.join(root, "Pending Pickup Template.xlsm")
test_copy = os.path.join(root, "scripts", "temp", "refs.xlsm")

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
try:
    wb = excel.Workbooks.Open(test_copy)
    print("References:")
    for ref in wb.VBProject.References:
        print(f"  name={ref.Name!r} major={ref.Major} minor={ref.Minor} builtin={ref.BuiltIn} "
              f"guid={ref.Guid!r} full={ref.FullPath!r}")
    print("\nComponents:")
    for comp in wb.VBProject.VBComponents:
        cm = comp.CodeModule
        print(f"  {comp.Name}: type={comp.Type} lines={cm.CountOfLines}")
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
