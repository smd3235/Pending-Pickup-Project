import os
import shutil

import psutil
import pythoncom
import win32com.client

root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
template = os.path.join(root, "Pending Pickup Template.xlsm")
test_copy = os.path.join(root, "scripts", "temp", "names_check.xlsm")

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
    print("Defined names:")
    for nm in wb.Names:
        print(f"  {nm.Name!r} -> {nm.RefersTo!r}")
    for sh in wb.Worksheets:
        try:
            for nm in sh.Names:
                print(f"  [{sh.Name}] {nm.Name!r} -> {nm.RefersTo!r}")
        except Exception:
            pass
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
