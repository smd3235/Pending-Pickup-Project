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
parser.add_argument("--skip-package", action="store_true", help="Skip rebuilding Pending Pickup Template (no-macros).xlsx and Project-Local-Build.zip")
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

    # Update START HERE labels/info (manager/director config in hidden template tabs)
    config_cells = {
        "C16": "Manager Email Template (hidden tab):",
        "D16": "Use 'Setup Manager Email Settings' to edit To/CC/Subject/Body.",
        "C17": "Director Email Template (hidden tab):",
        "D17": "Use 'Setup Director Email Template' to edit To/CC/Subject/Body.",
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

    def format_email_template_sheet(sh, title, field_values, hide=True):
        # field_values: dict with keys to, cc, subject, body
        try:
            if hide:
                sh.Visible = 0  # xlSheetHidden
        except Exception:
            pass
        try:
            sh.Cells.Clear
            sh.Cells.Interior.ColorIndex = 0
            sh.Cells.Font.Name = "Segoe UI"
            sh.Cells.Font.Color = 0x222222
            sh.Cells.Borders.LineStyle = -4142  # xlLineStyleNone
        except Exception:
            pass

        try:
            sh.Range("B2").Value = title
            sh.Range("B3").Value = "Fill in the fields below. Only these input areas are visible."
            sh.Range("B2").Font.Name = "Segoe UI"
            sh.Range("B2").Font.Bold = True
            sh.Range("B2").Font.Size = 20
            sh.Range("B2").Font.Color = 0x79411F  # dark blue-ish
            sh.Range("B3").Font.Name = "Segoe UI"
            sh.Range("B3").Font.Size = 10
            sh.Range("B3").Font.Color = 0x666666
        except Exception:
            pass

        fields = [
            ("C5", "TO:", "D5"),
            ("C6", "CC:", "D6"),
            ("C7", "SUBJECT:", "D7"),
        ]
        for label_cell, label_text, input_cell in fields:
            try:
                lc = sh.Range(label_cell)
                lc.Value = label_text
                lc.Font.Bold = True
                lc.Font.Size = 10
                lc.Font.Color = 0x555555
                lc.HorizontalAlignment = -4131  # xlLeft
                lc.VerticalAlignment = -4108   # xlCenter
            except Exception:
                pass
            try:
                rg = sh.Range(input_cell)
                rg.Interior.Color = 0xFFFFFF
                rg.Font.Name = "Segoe UI"
                rg.Font.Size = 11
                rg.HorizontalAlignment = -4131
                rg.VerticalAlignment = -4108
                rg.WrapText = False
                for edge in (7, 8, 9, 10):
                    try:
                        b = rg.Borders(edge)
                        b.LineStyle = 1
                        b.Weight = 2
                        b.Color = 0xD0D7DE
                    except Exception:
                        pass
            except Exception:
                pass

        try:
            sh.Range("C8").Value = "BODY:"
            sh.Range("C8").Font.Bold = True
            sh.Range("C8").Font.Size = 10
            sh.Range("C8").Font.Color = 0x555555
            sh.Range("C8").HorizontalAlignment = -4131
            sh.Range("C8").VerticalAlignment = -4108
        except Exception:
            pass
        try:
            body_rg = sh.Range("D9:K24")
            body_rg.UnMerge()
            body_rg.Merge()
            body_rg.Interior.Color = 0xFFFFFF
            body_rg.Font.Name = "Segoe UI"
            body_rg.Font.Size = 11
            body_rg.HorizontalAlignment = -4131
            body_rg.VerticalAlignment = -4160  # xlTop
            body_rg.WrapText = True
            for edge in (7, 8, 9, 10):
                try:
                    b = body_rg.Borders(edge)
                    b.LineStyle = 1
                    b.Weight = 2
                    b.Color = 0xD0D7DE
                except Exception:
                    pass
        except Exception:
            pass

        # Set values
        try:
            sh.Range("D5").Value = field_values.get("to", "")
            sh.Range("D6").Value = field_values.get("cc", "")
            sh.Range("D7").Value = field_values.get("subject", "")
            sh.Range("D9").Value = field_values.get("body", "")
        except Exception:
            pass

        # Layout sizing
        try:
            sh.Columns("A:B").ColumnWidth = 2
            sh.Columns("C").ColumnWidth = 13
            sh.Columns("D:K").ColumnWidth = 22
            sh.Rows("1").RowHeight = 10
            sh.Rows("2").RowHeight = 30
            sh.Rows("3").RowHeight = 18
            sh.Rows("4").RowHeight = 6
            for r in ("5", "6", "7"):
                sh.Rows(r).RowHeight = 28
            sh.Rows("8").RowHeight = 24
            sh.Rows("9").RowHeight = 26
            for rr in range(10, 25):
                sh.Rows(str(rr)).RowHeight = 22
        except Exception:
            pass

        # Hide non-input areas for clean look
        try:
            sh.Rows("25:1048576").Hidden = True
            sh.Columns("L:XFD").Hidden = True
            sh.Columns("A").Hidden = True
            sh.Columns("B").Hidden = True
        except Exception:
            pass

        # Gridlines/headings off in active window context best-effort
        try:
            aw = wb.Application.ActiveWindow
            if aw is not None:
                aw.DisplayGridlines = False
                aw.DisplayHeadings = False
        except Exception:
            pass

    def get_or_create_sheet(name):
        for w in wb.Worksheets:
            try:
                if w.Name == name:
                    return w
            except Exception:
                pass
        sh = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        sh.Name = name
        return sh

    # Ensure hidden Manager Email Template sheet exists with TO/CC/SUBJECT/BODY config.
    mgr_ws = get_or_create_sheet("Manager Email Template")
    try:
        # migrate old values if present
        mgr_to = str(mgr_ws.Range("D5").Value or "").strip() if str(mgr_ws.Range("C5").Value or "").upper().startswith("TO") else ""
        if not mgr_to:
            mgr_to = str(mgr_ws.Range("D4").Value or "").strip()
        mgr_cc = str(mgr_ws.Range("D6").Value or mgr_ws.Range("D5").Value or "").strip()
        mgr_subj = str(mgr_ws.Range("D7").Value or mgr_ws.Range("D6").Value or "").strip() or "Pending Pickup Report - Action Required"
        mgr_body = str(mgr_ws.Range("D9").Value or mgr_ws.Range("D7").Value or "").strip() or "Please review the Pending Pickup report snapshot below."
        # if layout was old (CC at D5), avoid duplicating
        if str(mgr_ws.Range("C5").Value or "").upper().startswith("CC"):
            mgr_cc = str(mgr_ws.Range("D5").Value or "").strip()
            mgr_subj = str(mgr_ws.Range("D6").Value or "").strip() or mgr_subj
            mgr_body = str(mgr_ws.Range("D7").Value or "").strip() or mgr_body
            mgr_to = ""
    except Exception:
        mgr_to, mgr_cc, mgr_subj, mgr_body = "", "", "Pending Pickup Report - Action Required", "Please review the Pending Pickup report snapshot below."
    format_email_template_sheet(
        mgr_ws,
        "Manager Email Template",
        {"to": mgr_to, "cc": mgr_cc, "subject": mgr_subj, "body": mgr_body},
        hide=True,
    )

    # Ensure hidden Director Email Template sheet exists with TO/CC/SUBJECT/BODY config.
    director_ws = get_or_create_sheet("Director Email Template")
    try:
        d_to = str(director_ws.Range("D5").Value or "").strip() if str(director_ws.Range("C5").Value or "").upper().startswith("TO") else str(director_ws.Range("D5").Value or "").strip()
        d_cc = str(director_ws.Range("D6").Value or "").strip()
        d_subj = str(director_ws.Range("D7").Value or "").strip() or "Pending Pickup Report - Daily Dashboard"
        d_body = str(director_ws.Range("D9").Value or director_ws.Range("D8").Value or "").strip() or "Please find the Pending Pickup daily dashboard below."
    except Exception:
        d_to, d_cc, d_subj, d_body = "", "", "Pending Pickup Report - Daily Dashboard", "Please find the Pending Pickup daily dashboard below."
    format_email_template_sheet(
        director_ws,
        "Director Email Template",
        {"to": d_to, "cc": d_cc, "subject": d_subj, "body": d_body},
        hide=True,
    )

    wb.Save()
    wb.Close(True)
    print(f"Synced {len(bas_files)} module(s) and button '{button_name}' into {template}")
    for b in bas_files:
        print(f"  - {os.path.basename(b)}")
finally:
    excel.Quit()

if not args.skip_package:
    no_macro_path = os.path.join(project_root, "Pending Pickup Template (no-macros).xlsx")
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
