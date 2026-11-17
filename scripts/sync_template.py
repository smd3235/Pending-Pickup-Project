# sync_template.py
# Re-imports all .bas VBA modules into the template's embedded VBA project,
# ensures START HERE buttons are installed, and rebuilds the work-PC local
# build package (macro-free template + module sources + instructions ZIP).
# Run this after every change to any .bas/.cls/.frm file.

import argparse
import glob
import os
import sys
import zipfile

import win32com.client

parser = argparse.ArgumentParser(description="Sync VBA macro modules into the template and rebuild local build package")
parser.add_argument("--template", default=None)
parser.add_argument("--bas-dir", default=None, help="Directory containing .bas modules (default: project root)")
parser.add_argument("--bas", action="append", default=[], help="Specific .bas file to sync (repeatable)")
parser.add_argument("--button", default="btnCreatePendingPickupReport")
parser.add_argument("--skip-package", action="store_true", help="Skip rebuilding Template (No Macros).xlsx and Project-Local-Build.zip")
args = parser.parse_args()

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)

template = args.template or os.path.join(project_root, "Pending Pickup Template.xlsm")
bas_dir = args.bas_dir or project_root
button_name = args.button

# Collect .bas files to sync: explicit --bas first, then all .bas in bas_dir
bas_files = list(args.bas)
if not bas_files:
    all_bas = sorted(glob.glob(os.path.join(bas_dir, "*.bas")))
    # Exclude the old monolithic file; it is replaced by the modular .bas files.
    bas_files = [b for b in all_bas if os.path.basename(b) != "Pending_Pickup_Macro.bas"]

if not os.path.exists(template):
    print(f"Template not found: {template}", file=sys.stderr)
    sys.exit(1)

for b in bas_files:
    if not os.path.exists(b):
        print(f"Bas not found: {b}", file=sys.stderr)
        sys.exit(1)

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable

try:
    wb = excel.Workbooks.Open(template)
    vbp = wb.VBProject

    # Track which module names we are importing so we can remove stale ones.
    imported_names = set()
    for bas in bas_files:
        module_name = os.path.splitext(os.path.basename(bas))[0]
        imported_names.add(module_name)

    # Remove existing components whose names collide with imported modules.
# Also remove any leftover PendingPickupReport* modules from previous syncs.
    removal_names = imported_names
    for c in list(vbp.VBComponents):
        if c.Name in removal_names or c.Name.startswith("PendingPickupReport"):
            vbp.VBComponents.Remove(c)

    # Import each .bas module.
    for bas in bas_files:
        vbp.VBComponents.Import(bas)

    # Ensure the styled button shape exists and points at the macro.
    ws = wb.Worksheets("START HERE")
    btn = None
    for s in ws.Shapes:
        if s.Name == button_name:
            btn = s
            break
    if btn is None:
        cta = ws.Range("B14")
        btn = ws.Shapes.AddShape(5, cta.Left + 20, cta.Top + 8, 360, 48)
        btn.Name = button_name
        btn.Fill.ForeColor.RGB = 7949855
        btn.Line.ForeColor.RGB = 4116214
        btn.TextFrame2.TextRange.Text = "CREATE PENDING PICKUP REPORT"
        btn.TextFrame2.TextRange.Font.Bold = True
        btn.TextFrame2.TextRange.Font.Size = 14
        btn.TextFrame2.TextRange.Font.Name = "Segoe UI"
        btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
    btn.OnAction = "CreatePendingPickupReport"

    # Ensure the Setup Email CC button shape exists on START HERE.
    setup_btn = None
    for s in ws.Shapes:
        if s.Name == "btnSetupEmailCC":
            setup_btn = s
            break
    if setup_btn is None:
        cta = ws.Range("B14")
        setup_btn = ws.Shapes.AddShape(5, cta.Left + 390, cta.Top + 8, 200, 48)
        setup_btn.Name = "btnSetupEmailCC"
        setup_btn.Fill.ForeColor.RGB = 7949855
        setup_btn.Line.ForeColor.RGB = 4116214
        setup_btn.TextFrame2.TextRange.Text = "Setup Manager Email Settings"
        setup_btn.TextFrame2.TextRange.Font.Bold = True
        setup_btn.TextFrame2.TextRange.Font.Size = 11
        setup_btn.TextFrame2.TextRange.Font.Name = "Segoe UI"
        setup_btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
    else:
        try:
            setup_btn.TextFrame2.TextRange.Text = "Setup Manager Email Settings"
        except Exception:
            pass
    setup_btn.OnAction = "SetupEmailCC"

    # Ensure the Setup Director Email button shape exists on START HERE.
    director_btn = None
    for s in ws.Shapes:
        if s.Name == "btnSetupDirectorEmail":
            director_btn = s
            break
    if director_btn is None:
        cta = ws.Range("B14")
        director_btn = ws.Shapes.AddShape(5, cta.Left + 600, cta.Top + 8, 220, 48)
        director_btn.Name = "btnSetupDirectorEmail"
        director_btn.Fill.ForeColor.RGB = 7949855
        director_btn.Line.ForeColor.RGB = 4116214
        director_btn.TextFrame2.TextRange.Text = "Setup Director Email Template"
        director_btn.TextFrame2.TextRange.Font.Bold = True
        director_btn.TextFrame2.TextRange.Font.Size = 11
        director_btn.TextFrame2.TextRange.Font.Name = "Segoe UI"
        director_btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
    else:
        try:
            director_btn.TextFrame2.TextRange.Text = "Setup Director Email Template"
        except Exception:
            pass
    director_btn.OnAction = "SetupDirectorEmail"

    # Update START HERE labels/info (manager config migrated to hidden Manager Email Template tab)
    config_cells = {
        "C16": "Manager Email Template (hidden tab):",
        "D16": "Use 'Setup Manager Email Settings' to edit CC/Subject/Intro.",
        "C17": "Director Email Template (hidden tab):",
        "D17": "Use 'Setup Director Email Template' to edit To/CC/Subject/Intro.",
        "C18": "",
        "D18": "",
        "C20": "",
        "D20": "",
    }
    for addr, val in config_cells.items():
        try:
            ws.Range(addr).Value = val
        except Exception:
            pass

    # Style the label cells
    for label_addr in ["C16", "C17"]:
        try:
            ws.Range(label_addr).Font.Bold = True
            ws.Range(label_addr).Font.Size = 10
        except Exception:
            pass
    for label_addr in ["C18", "C20"]:
        try:
            ws.Range(label_addr).Font.Bold = False
        except Exception:
            pass

    def ensure_template_sheet(sheet_name, title, subtitle, rows, hide=True):
        sh = None
        for w in wb.Worksheets:
            try:
                if w.Name == sheet_name:
                    sh = w
                    break
            except Exception:
                pass
        if sh is None:
            sh = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            sh.Name = sheet_name
        try:
            if hide:
                sh.Visible = 0  # xlSheetHidden
        except Exception:
            pass
        cells = {"B2": title, "B3": subtitle}
        cells.update(rows)
        for addr, val in cells.items():
            try:
                existing = sh.Range(addr).Value
                if existing is None or str(existing).strip() == "" or addr in ["B2", "B3"] or addr.startswith("C"):
                    sh.Range(addr).Value = val
            except Exception:
                pass
        try:
            sh.Range("B2").Font.Bold = True
            sh.Range("B2").Font.Size = 14
            sh.Range("B3").Font.Size = 10
            for r in rows.keys():
                if r.startswith("C"):
                    sh.Range(r).Font.Bold = True
            sh.Columns("C").ColumnWidth = 34
            sh.Columns("D").ColumnWidth = 60
        except Exception:
            pass
        return sh

    # Ensure hidden Manager Email Template sheet exists with CC/Subject/Intro config.
    mgr_ws = ensure_template_sheet(
        "Manager Email Template",
        "Manager Email Template",
        "Edit the values in column D. This tab can remain hidden; manager email buttons read from it.",
        {
            "C5": "CC Recipients (comma-separated):",
            "D5": ws.Range("D16").Value if str(ws.Range("C16").Value or "").startswith("Email CC") else "",
            "C6": "Subject (leave blank for default):",
            "D6": "Pending Pickup Report - Action Required",
            "C7": "Intro message (leave blank for default):",
            "D7": "Please review the Pending Pickup report snapshot below.",
        },
        hide=True,
    )
    try:
        mgr_ws.Range("D7").WrapText = True
    except Exception:
        pass

    # Ensure hidden Director Email Template sheet exists with To/CC/Subject/Body config.
    director_ws = ensure_template_sheet(
        "Director Email Template",
        "Director / Dashboard Email Template",
        "Edit the values in column D. This tab can remain hidden; the Dashboard email button reads from it.",
        {
            "C5": "To (comma-separated):",
            "D5": "",
            "C6": "CC (comma-separated):",
            "D6": "",
            "C7": "Subject (leave blank for default):",
            "D7": "Pending Pickup Report - Daily Dashboard",
            "C8": "Intro message (leave blank for default):",
            "D8": "Please find the Pending Pickup daily dashboard below.",
        },
        hide=True,
    )
    try:
        director_ws.Range("D8").WrapText = True
    except Exception:
        pass

    wb.Save()
    wb.Close(True)
    print(f"Synced {len(bas_files)} module(s) and button '{button_name}' into {template}")
    for b in bas_files:
        print(f"  - {os.path.basename(b)}")
finally:
    excel.Quit()

if not args.skip_package:
    no_macro_path = os.path.join(project_root, "Template (No Macros).xlsx")
    instructions_path = os.path.join(project_root, "BUILD-INSTRUCTIONS.txt")
    zip_path = os.path.join(project_root, "Project-Local-Build.zip")

    # Collect all VBA source modules in project root (.bas, .cls, .frm)
    module_files = []
    for ext in ("*.bas", "*.cls", "*.frm"):
        module_files.extend(glob.glob(os.path.join(project_root, ext)))
    module_files = sorted(set(module_files))
    if not module_files:
        print("No VBA module files (.bas/.cls/.frm) found for build package", file=sys.stderr)
        sys.exit(1)
    if not os.path.exists(instructions_path):
        print(f"Build instructions not found: {instructions_path}", file=sys.stderr)
        sys.exit(1)

    # Create macro-free base workbook from the synced template
    excel2 = win32com.client.Dispatch("Excel.Application")
    excel2.Visible = False
    excel2.DisplayAlerts = False
    excel2.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
    try:
        wb2 = excel2.Workbooks.Open(template)
        if os.path.exists(no_macro_path):
            os.remove(no_macro_path)
        # 51 = xlOpenXMLWorkbook (.xlsx, no macros)
        wb2.SaveAs(no_macro_path, FileFormat=51)
        wb2.Close(False)
        print(f"Created macro-free base: {no_macro_path}")
    finally:
        excel2.Quit()

    # Build ZIP with macro-free template, all modules, and instructions
    zip_items = [no_macro_path] + module_files + [instructions_path]
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for item in zip_items:
            zf.write(item, arcname=os.path.basename(item))
    print(f"Built local build package: {zip_path}")
    print("  Includes:")
    for item in zip_items:
        print(f"    - {os.path.basename(item)}")
