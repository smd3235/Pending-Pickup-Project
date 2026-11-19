# snapshot-report.py
# Exports every sheet (Dashboard, manager tabs, Raw Data) of a generated
# Pending Pickup report as a PNG into documentation\scripts\screenshots\.
# Uses CopyPicture -> standalone Chart sheet -> Chart.Export so whole sheets are
# captured as images for the README. Defaults to the newest workbook in Reports\.

import argparse
import os
import re
import subprocess
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
parser.add_argument("--template-path", default=None, help="Template workbook for START HERE sheet screenshot (defaults to Pending Pickup Template.xlsm)")
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
            if f.lower().endswith((".xlsx", ".xlsm")):
                candidates.append(os.path.join(reports_dir, f))
    if candidates:
        candidates.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        path = candidates[0]
        print(f"Using newest report: {path}")
    else:
        print("No workbook given and Reports\\ holds no .xlsx/.xlsm. Pass --path.", file=sys.stderr)
        sys.exit(2)

if not os.path.exists(path):
    print(f"Workbook not found: {path}", file=sys.stderr)
    sys.exit(2)

path = os.path.abspath(path)

os.makedirs(out_folder, exist_ok=True)

# Export screenshots for the selected sheets. Existing screenshots for this
# prefix are replaced only after a successful non-empty export (validated below),
# so a failed/empty export does not leave stale-corrupted finals without error.


def _export_chart_to_png(chart, png_path, label):
    tmp_png = png_path + ".tmp"
    last_err = None
    for attempt in range(3):
        last_err = None
        try:
            if os.path.exists(tmp_png):
                os.remove(tmp_png)
        except Exception:
            pass
        try:
            chart.Export(tmp_png)
        except Exception as e:
            last_err = e
            time.sleep(1.0)
            continue
        try:
            if os.path.exists(tmp_png) and os.path.getsize(tmp_png) > 0:
                os.replace(tmp_png, png_path)
                return
            last_err = RuntimeError(f"Export produced empty file for {label}")
        except Exception as e:
            last_err = e
        time.sleep(1.0)
    try:
        if os.path.exists(tmp_png):
            os.remove(tmp_png)
    except Exception:
        pass
    raise last_err or RuntimeError(f"Export failed for {label}")


def save_sheet_image(workbook, worksheet, png_path):
    used = worksheet.UsedRange
    first_row = used.Row
    first_col = used.Column
    last_row = first_row + used.Rows.Count - 1
    last_col = first_col + used.Columns.Count - 1

    # Include embedded charts/shapes so content below UsedRange (e.g. dashboard chart) is captured.
    try:
        for co in worksheet.ChartObjects():
            br = co.BottomRightCell
            tr = co.TopLeftCell
            if br.Row > last_row:
                last_row = br.Row
            if br.Column > last_col:
                last_col = br.Column
            if tr.Row < first_row:
                first_row = tr.Row
            if tr.Column < first_col:
                first_col = tr.Column
    except Exception:
        pass
    try:
        for shp in worksheet.Shapes:
            br = shp.BottomRightCell
            tr = shp.TopLeftCell
            if br.Row > last_row:
                last_row = br.Row
            if br.Column > last_col:
                last_col = br.Column
            if tr.Row < first_row:
                first_row = tr.Row
            if tr.Column < first_col:
                first_col = tr.Column
    except Exception:
        pass

    rng = worksheet.Range(worksheet.Cells(first_row, first_col), worksheet.Cells(last_row, last_col))

    # Activate sheet and set zoom so CopyPicture captures a stable, fully-rendered view.
    worksheet.Activate()
    try:
        workbook.Application.CutCopyMode = False
    except Exception:
        pass
    try:
        worksheet.Range("A1").Select()
    except Exception:
        pass
    try:
        workbook.Application.ActiveWindow.ScrollRow = 1
        workbook.Application.ActiveWindow.ScrollColumn = 1
        workbook.Application.ActiveWindow.Zoom = 100
    except Exception:
        pass
    try:
        workbook.Application.ScreenUpdating = True
    except Exception:
        pass
    try:
        workbook.Application.Calculate()
    except Exception:
        pass
    # Allow Excel time to render/layout (hidden sheets, charts, shapes, formatting).
    time.sleep(2.0)

    # Copy the range as a picture. Bitmap (Format=2) is more faithful to on-screen
    # rendering and avoids metafile artifacts that can expose hidden/overlay elements.
    copied = False
    for attempt in range(4):
        try:
            rng.CopyPicture(Appearance=1, Format=2)  # xlScreen=1, xlBitmap=2
            copied = True
            break
        except Exception:
            time.sleep(1.2)
    if not copied:
        # Fallback: picture/metafile
        for attempt in range(2):
            try:
                rng.CopyPicture(Appearance=1, Format=1)  # xlScreen=1, xlPicture=1
                copied = True
                break
            except Exception:
                time.sleep(1.2)
    if not copied:
        raise RuntimeError(f"CopyPicture failed after retries for {worksheet.Name}")

    chart = workbook.Charts.Add()
    try:
        pasted = False
        last_err = None
        for attempt in range(3):
            try:
                chart.Paste()
                pasted = True
                break
            except Exception as e:
                last_err = e
                time.sleep(0.8)
        if not pasted:
            raise RuntimeError(f"Paste failed for {worksheet.Name}: {last_err}")

        time.sleep(1.0)

        # Export to a temp .png file, validate non-empty, then replace target.
        base, ext = os.path.splitext(png_path)
        tmp_png = base + ".__tmp__" + (ext or ".png")
        last_err = None
        for attempt in range(3):
            last_err = None
            try:
                if os.path.exists(tmp_png):
                    os.remove(tmp_png)
            except Exception:
                pass
            try:
                chart.Export(tmp_png)
            except Exception as e:
                last_err = e
                time.sleep(1.0)
                continue
            try:
                if os.path.exists(tmp_png) and os.path.getsize(tmp_png) > 0:
                    os.replace(tmp_png, png_path)
                    return
                last_err = RuntimeError(f"Export produced empty file for {worksheet.Name}")
            except Exception as e:
                last_err = e
            time.sleep(1.0)

        try:
            if os.path.exists(tmp_png):
                os.remove(tmp_png)
        except Exception:
            pass
        raise last_err or RuntimeError(f"Export failed for {worksheet.Name}")
    finally:
        try:
            chart.Delete()
        except Exception:
            pass


# Kill any lingering Excel processes before starting (unless --no-kill is set,
# e.g. when called from run-report.py which still has Excel running).
if not args.no_kill:
    try:
        subprocess.run(["taskkill", "/F", "/IM", "excel.exe"], capture_output=True, timeout=10)
    except Exception:
        pass
    time.sleep(2.0)

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
    wb = xl.Workbooks.Open(path, False, False)
    time.sleep(3.0)

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

    # Also export START HERE sheet from the template workbook so users can see
    # the template layout (START HERE sheet with buttons).
    template_path = args.template_path or os.path.join(project_root, "Pending Pickup Template.xlsm")
    if os.path.exists(template_path):
        try:
            abs_tpl = os.path.abspath(template_path)
            wb2 = xl.Workbooks.Open(abs_tpl, False, False)
            time.sleep(3.0)
            try:
                start_here = None
                for s in wb2.Worksheets:
                    if s.Name == "START HERE":
                        start_here = s
                        break
                if start_here:
                    png = os.path.join(out_folder, "template-start-here.png")
                    save_sheet_image(wb2, start_here, png)
                    kb = round(os.path.getsize(png) / 1024, 1)
                    print(f"  exported {png}  ({kb} KB)")
                    exported += 1
                else:
                    print("START HERE sheet not found in template", file=sys.stderr)
            except Exception as e:
                print(f"Template START HERE screenshot failed: {e}", file=sys.stderr)
            finally:
                # Also export email template sheets from the template workbook
                # so users can see the manager/director email setup layouts.
                email_sheet_names = ["Manager Email Template", "Director Email Template"]
                for email_name in email_sheet_names:
                    try:
                        found = None
                        for s in wb2.Worksheets:
                            if s.Name == email_name:
                                found = s
                                break
                        if found:
                            safe = re.sub(r"[^\w\-]+", "-", email_name).strip("-").lower()
                            png = os.path.join(out_folder, f"template-{safe}.png")
                            save_sheet_image(wb2, found, png)
                            kb = round(os.path.getsize(png) / 1024, 1)
                            print(f"  exported {png}  ({kb} KB)")
                            exported += 1
                        else:
                            print(f"Email template sheet not found: {email_name}", file=sys.stderr)
                    except Exception as e:
                        print(f"Template email screenshot ({email_name}) failed: {e}", file=sys.stderr)
                wb2.Close(False)
        except Exception as e:
            print(f"Template screenshot failed: {e}", file=sys.stderr)
finally:
    if wb:
        wb.Close(False)
    xl.Quit()
    try:
        subprocess.run(["taskkill", "/F", "/IM", "excel.exe"], capture_output=True, timeout=10)
    except Exception:
        pass

if exported == 0 and not args.list and args.sheets:
    print("Nothing exported.", file=sys.stderr)
    sys.exit(3)

sys.exit(0)
