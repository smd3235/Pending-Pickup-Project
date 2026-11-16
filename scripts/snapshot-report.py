# snapshot-report.py
# Exports every sheet (Dashboard, manager tabs, Raw Data) of a generated
# Pending Pickup report as a PNG into documentation\scripts\screenshots\.
# Uses CopyPicture -> standalone Chart sheet -> Chart.Export so whole sheets are
# captured as images for the README. Defaults to the newest workbook in Reports\.

import argparse
import os
import re
import sys
import time

import win32com.client

parser = argparse.ArgumentParser(description="Export report sheets as PNG screenshots")
parser.add_argument("--path", default=None, help="Workbook to export (defaults to newest in Reports/)")
parser.add_argument("--sheets", nargs="*", default=None, help="Specific sheet names to export")
parser.add_argument("--out", default=None, help="Output folder (defaults to documentation/screenshots/)")
parser.add_argument("--prefix", default="report", help="Filename prefix")
parser.add_argument("--list", action="store_true", help="List sheets and exit")
parser.add_argument("--no-kill", action="store_true", help="Do not kill Excel processes before starting (used when called from run-report.py)")
args = parser.parse_args()

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)

path = args.path
out_folder = args.out or os.path.join(project_root, "documentation", "screenshots")
prefix = args.prefix

if not path:
    reports_dir = os.path.join(project_root, "Reports")
    candidates = []
    if os.path.isdir(reports_dir):
        for f in os.listdir(reports_dir):
            if f.lower().endswith(".xlsx"):
                candidates.append(os.path.join(reports_dir, f))
    if candidates:
        candidates.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        path = candidates[0]
        print(f"Using newest report: {path}")
    else:
        print("No workbook given and Reports\\ holds no .xlsx. Pass --path.", file=sys.stderr)
        sys.exit(2)

if not os.path.exists(path):
    print(f"Workbook not found: {path}", file=sys.stderr)
    sys.exit(2)

path = os.path.abspath(path)

os.makedirs(out_folder, exist_ok=True)

# Always start fresh: delete every existing PNG so committed screenshots
# reflect only the latest report run, never a stale mix of old/new sheets.
for f in os.listdir(out_folder):
    if f.lower().endswith(".png"):
        try:
            os.remove(os.path.join(out_folder, f))
        except OSError:
            pass


def save_sheet_image(workbook, worksheet, png_path):
    used = worksheet.UsedRange
    first_row = used.Row
    first_col = used.Column
    last_row = first_row + used.Rows.Count - 1
    last_col = first_col + used.Columns.Count - 1
    rng = worksheet.Range(worksheet.Cells(first_row, first_col), worksheet.Cells(last_row, last_col))

    # Activate sheet and set zoom to 100% so CopyPicture captures at full resolution,
    # not zoomed out to 85% (which makes everything look squished).
    worksheet.Activate()
    workbook.Application.ActiveWindow.Zoom = 100

    rng.CopyPicture(1, 2)  # xlScreen=1, xlBitmap=2

    rows = int(used.Rows.Count)
    cols = int(used.Columns.Count)
    # Wider chart so content isn't squished onto one page width.
    w = max(1800, cols * 110)
    h = max(900, rows * 30)

    chart = workbook.Charts.Add()
    try:
        chart.Paste()
        time.sleep(0.25)
        if os.path.exists(png_path):
            os.remove(png_path)
        chart.Export(png_path)
    finally:
        chart.Delete()


# Kill any lingering Excel processes before starting (unless --no-kill is set,
# e.g. when called from run-report.py which still has Excel running).
import subprocess
if not args.no_kill:
    try:
        subprocess.run(["taskkill", "/F", "/IM", "excel.exe"], capture_output=True)
    except Exception:
        pass
    time.sleep(3.0)

xl = win32com.client.Dispatch("Excel.Application")
xl.Visible = True
xl.WindowState = -4137  # xlMaximized
xl.DisplayAlerts = False
xl.AutomationSecurity = 3
# Reset Excel state so CopyPicture works reliably (non-critical if it fails).
try:
    xl.ScreenUpdating = True
    xl.Calculation = -4135  # xlCalculationAutomatic
    xl.EnableEvents = True
except Exception:
    pass
wb = None
exported = 0
try:
    wb = xl.Workbooks.Open(path, False, True)
    time.sleep(1.0)

    if args.list or not args.sheets:
        print(f"Sheets in {os.path.basename(path)}:")
        for s in wb.Worksheets:
            u = s.UsedRange
            addr = u.Address
            print(f"  {s.Name:<28} used {addr}")

    targets = []
    if args.sheets:
        for name in args.sheets:
            found = None
            for s in wb.Worksheets:
                if s.Name == name:
                    found = s
                    break
            if found:
                targets.append(found)
            else:
                print(f"Sheet not found, skipped: {name}", file=sys.stderr)
    else:
        targets = list(wb.Worksheets)

    for ws in targets:
        safe = re.sub(r"[^\w\-]+", "-", ws.Name).strip("-").lower()
        png = os.path.join(out_folder, f"{prefix}-{safe}.png")
        save_sheet_image(wb, ws, png)
        kb = round(os.path.getsize(png) / 1024, 1)
        print(f"  exported {png}  ({kb} KB)")
        exported += 1
finally:
    if wb:
        wb.Close(False)
    xl.Quit()
    try:
        subprocess.run(["taskkill", "/F", "/IM", "excel.exe"], capture_output=True)
    except Exception:
        pass

if exported == 0 and not args.list and args.sheets:
    print("Nothing exported.", file=sys.stderr)
    sys.exit(3)

sys.exit(0)
