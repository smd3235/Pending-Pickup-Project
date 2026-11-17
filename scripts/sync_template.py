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

    # Ensure START HERE action buttons exist and are aligned cleanly.
    ws = wb.Worksheets("START HERE")
    cta = ws.Range("B14")

    # Reserve vertical space for a 3-button stack so text below doesn't overlap
    try:
        ws.Rows("14").RowHeight = 165
        ws.Rows("15").RowHeight = 8
        ws.Rows("16").RowHeight = 8
        ws.Rows("17").RowHeight = 8
    except Exception:
        pass

    row1_top = cta.Top + 6
    row2_top = row1_top + 52
    row3_top = row2_top + 52
    btn_left = cta.Left + 12
    btn_w = 360
    btn_h = 44

    btn = None
    for s in ws.Shapes:
        if s.Name == button_name:
            btn = s
            break
    if btn is None:
        btn = ws.Shapes.AddShape(5, btn_left, row1_top, btn_w, btn_h)
        btn.Name = button_name
        btn.Fill.ForeColor.RGB = 7949855
        btn.Line.ForeColor.RGB = 4116214
        btn.TextFrame2.TextRange.Font.Bold = True
        btn.TextFrame2.TextRange.Font.Size = 14
        btn.TextFrame2.TextRange.Font.Name = "Segoe UI"
        btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
    btn.OnAction = "CreatePendingPickupReport"
    try:
        btn.Left = btn_left
        btn.Top = row1_top
        btn.Width = btn_w
        btn.Height = btn_h
        btn.TextFrame2.TextRange.Text = "CREATE PENDING PICKUP REPORT"
    except Exception:
        pass

    # Setup Manager Email button
    setup_btn = None
    for s in ws.Shapes:
        if s.Name == "btnSetupEmailCC":
            setup_btn = s
            break
    if setup_btn is None:
        setup_btn = ws.Shapes.AddShape(5, btn_left, row2_top, 280, btn_h)
        setup_btn.Name = "btnSetupEmailCC"
        setup_btn.Fill.ForeColor.RGB = 7949855
        setup_btn.Line.ForeColor.RGB = 4116214
        setup_btn.TextFrame2.TextRange.Font.Bold = True
        setup_btn.TextFrame2.TextRange.Font.Size = 11
        setup_btn.TextFrame2.TextRange.Font.Name = "Segoe UI"
        setup_btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
    setup_btn.OnAction = "SetupEmailCC"
    try:
        setup_btn.Left = btn_left
        setup_btn.Top = row2_top
        setup_btn.Width = 280
        setup_btn.Height = btn_h
        setup_btn.TextFrame2.TextRange.Text = "Setup Manager Email Settings"
    except Exception:
        pass

    # Setup Director Email button
    director_btn = None
    for s in ws.Shapes:
        if s.Name == "btnSetupDirectorEmail":
            director_btn = s
            break
    if director_btn is None:
        director_btn = ws.Shapes.AddShape(5, btn_left, row3_top, 280, btn_h)
        director_btn.Name = "btnSetupDirectorEmail"
        director_btn.Fill.ForeColor.RGB = 7949855
        director_btn.Line.ForeColor.RGB = 4116214
        director_btn.TextFrame2.TextRange.Font.Bold = True
        director_btn.TextFrame2.TextRange.Font.Size = 11
        director_btn.TextFrame2.TextRange.Font.Name = "Segoe UI"
        director_btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 16777215
    director_btn.OnAction = "SetupDirectorEmail"
    try:
        director_btn.Left = btn_left
        director_btn.Top = row3_top
        director_btn.Width = 280
        director_btn.Height = btn_h
        director_btn.TextFrame2.TextRange.Text = "Setup Director Email Template"
    except Exception:
        pass

    # Reorganize START HERE text so buttons do not overlap content
    start_here_text = {
        "B16": "",
        "B17": "",
        "B18": "",
        "B19": "",
        "B20": "",
        "B21": "",
        "B22": "",
        "B23": "",
        "B18": "GOOD TO KNOW",
        "B19": "•  Only rows with Treatment Status 'Pending Pickup' are included.",
        "B20": "•  Pickup Age is days since AUDIT: TP End Date — older means higher priority.",
        "B21": "•  A missing AUDIT: TP End Date shows as 'MISSING DATE'.",
        "B23": "v1.0  •  Pending Pickup Report Builder",
    }
    # apply in stable order (clear/move)
    ordered_updates = [
        ("B16", ""), ("B17", ""),
        ("B18", "GOOD TO KNOW"),
        ("B19", "•  Only rows with Treatment Status 'Pending Pickup' are included."),
        ("B20", "•  Pickup Age is days since AUDIT: TP End Date — older means higher priority."),
        ("B21", "•  A missing AUDIT: TP End Date shows as 'MISSING DATE'."),
        ("B22", ""),
        ("B23", "v1.0  •  Pending Pickup Report Builder"),
    ]
    for addr, val in ordered_updates:
        try:
            ws.Range(addr).Value = val
        except Exception:
            pass
    try:
        ws.Rows("18").RowHeight = 24
        ws.Rows("19").RowHeight = 20.1
        ws.Rows("20").RowHeight = 20.1
        ws.Rows("21").RowHeight = 20.1
        ws.Rows("22").RowHeight = 8
        ws.Rows("23").RowHeight = 18
        ws.Range("B18").Font.Bold = True
        ws.Range("B18").Font.Size = 11
    except Exception:
        pass

    # Update START HERE labels/info (manager/director config in hidden template tabs)
    config_cells = {
        "C16": "",
        "D16": "",
        "C17": "",
        "D17": "",
        "C18": "",
        "D18": "",
        "C19": "",
        "D19": "",
        "C24": "Manager Email Template (hidden tab):",
        "D24": "Use 'Setup Manager Email Settings' to edit To/CC/Subject/Body.",
        "C25": "Director Email Template (hidden tab):",
        "D25": "Use 'Setup Director Email Template' to edit To/CC/Subject/Body.",
        "C20": "",
        "D20": "",
    }
    for addr, val in config_cells.items():
        try:
            ws.Range(addr).Value = val
        except Exception:
            pass

    # Style the label cells
    for label_addr in ["C24", "C25"]:
        try:
            ws.Range(label_addr).Font.Bold = True
            ws.Range(label_addr).Font.Size = 10
        except Exception:
            pass

    def format_email_template_sheet(sh, title, field_values, hide=True, accent=0x79411F, accent_light=0xF2E8D5, badge_text="EMAIL TEMPLATE"):
        # field_values: dict with keys to, cc, subject, body
        # NOTE: VBA readers expect D5 (To), D6 (CC), D7 (Subject), D9 (Body) - keep these cells.
        try:
            if hide:
                sh.Visible = 0  # xlSheetHidden
        except Exception:
            pass

        # Reset sheet for a clean, shape-free canvas (no overlay shapes/text obstruction)
        try:
            for s in list(sh.Shapes):
                try:
                    s.Delete()
                except Exception:
                    pass
        except Exception:
            pass
        try:
            sh.Cells.ClearContents()
            sh.Cells.ClearFormats()
            sh.Cells.UnMerge()
            sh.Cells.Interior.Color = 0xF4F1EB  # warm canvas
            sh.Cells.Font.Name = "Segoe UI"
            sh.Cells.Font.Size = 11
            sh.Cells.Font.Color = 0x1F1F1F
            sh.Cells.Borders.LineStyle = -4142  # xlLineStyleNone
            sh.Cells.WrapText = False
            sh.Cells.RowHeight = 18
        except Exception:
            pass

        # Layout (preserve VBA-read cells): header 2-3, instructions 4, TO 5, gap 6 spacer row? no-can't shift.
        # Keep fixed cells D5/D6/D7/D9; use spacer rows 4 and 8 and visual card separation.
        try:
            sh.Columns("A").ColumnWidth = 1.0
            sh.Columns("B").ColumnWidth = 1.4
            sh.Columns("C").ColumnWidth = 12.5
            for col in ["D", "E", "F", "G", "H", "I", "J", "K"]:
                sh.Columns(col).ColumnWidth = 18.5
            sh.Rows("1").RowHeight = 6
            sh.Rows("2").RowHeight = 38
            sh.Rows("3").RowHeight = 18
            sh.Rows("4").RowHeight = 18   # instructions
            sh.Rows("5").RowHeight = 38   # TO (D5)
            sh.Rows("6").RowHeight = 38   # CC (D6)
            sh.Rows("7").RowHeight = 38   # SUBJECT (D7)
            sh.Rows("8").RowHeight = 16   # spacer before BODY
            sh.Rows("9").RowHeight = 26   # BODY top (merged D9:K30)
            for rr in range(10, 31):
                sh.Rows(str(rr)).RowHeight = 20
        except Exception:
            pass

        try:
            sh.Rows("31:1048576").Hidden = True
            sh.Columns("L:XFD").Hidden = True
            sh.Columns("A").Hidden = True
        except Exception:
            pass

        # Header band
        try:
            sh.Range("B2:K2").Merge()
            sh.Range("B3:K3").Merge()
            sh.Range("B2:K3").Interior.Color = accent
            sh.Range("B2").Value = title
            sh.Range("B3").Value = "Enter email details below. Report content is added automatically when you generate the email."
            sh.Range("B2").Font.Name = "Segoe UI Semibold"
            sh.Range("B2").Font.Bold = True
            sh.Range("B2").Font.Size = 19
            sh.Range("B2").Font.Color = 0xFFFFFF
            sh.Range("B2").HorizontalAlignment = -4131
            sh.Range("B2").VerticalAlignment = -4108
            sh.Range("B2").IndentLevel = 1
            sh.Range("B3").Font.Name = "Segoe UI"
            sh.Range("B3").Font.Size = 10
            sh.Range("B3").Font.Color = 0xF6F0E7
            sh.Range("B3").HorizontalAlignment = -4131
            sh.Range("B3").VerticalAlignment = -4108
            sh.Range("B3").IndentLevel = 1
            # Reserve top-right area for Return button (keep cells clean)
            try:
                sh.Range("J2:K2").UnMerge()
                sh.Range("J2:K2").Interior.Color = accent
                sh.Range("J2:K2").ClearContents()
            except Exception:
                pass
            for edge in (7, 8, 9, 10):
                try:
                    b = sh.Range("B2:K3").Borders(edge)
                    b.LineStyle = 1
                    b.Weight = 2
                    b.Color = accent
                except Exception:
                    pass
        except Exception:
            pass

        # Return to START HERE / re-hide button (shape in header only; does not cover input fields)
        try:
            btn_cell = sh.Range("J2:K2")
            btn = sh.Shapes.AddShape(5, btn_cell.Left + 4, btn_cell.Top + 5, btn_cell.Width - 8, 26)
            btn.Name = "btnReturnStartHere"
            btn.OnAction = "ReturnToStartHere"
            btn.Fill.ForeColor.RGB = accent_light
            btn.Line.ForeColor.RGB = 0xFFFFFF
            btn.Line.Weight = 1
            btn.TextFrame2.TextRange.Text = "Return to START HERE"
            btn.TextFrame2.TextRange.Font.Name = "Segoe UI"
            btn.TextFrame2.TextRange.Font.Bold = True
            btn.TextFrame2.TextRange.Font.Size = 9
            btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = accent
            btn.TextFrame2.TextRange.ParagraphFormat.Alignment = 2  # center
            btn.TextFrame2.VerticalAnchor = 3  # middle
            try:
                btn.Placement = 2  # move without sizing
            except Exception:
                pass
        except Exception:
            pass

        # Instructions row
        try:
            sh.Range("B4:K4").Merge()
            sh.Range("B4").Value = "TO / CC: enter email addresses separated by semicolons (;)  •  SUBJECT/BODY: leave blank to use defaults where available."
            sh.Range("B4").Font.Name = "Segoe UI"
            sh.Range("B4").Font.Size = 9
            sh.Range("B4").Font.Italic = True
            sh.Range("B4").Font.Color = 0x5A5348
            sh.Range("B4").Interior.Color = 0xEDE8DF
            sh.Range("B4").HorizontalAlignment = -4131
            sh.Range("B4").VerticalAlignment = -4108
            sh.Range("B4").IndentLevel = 1
            for edge in (8, 9):  # bottom/top subtle
                try:
                    b = sh.Range("B4:K4").Borders(edge)
                    b.LineStyle = 1
                    b.Weight = 2
                    b.Color = 0xDDD6CA
                except Exception:
                    pass
        except Exception:
            pass

        def _style_label(addr, text, fill):
            try:
                rg = sh.Range(addr)
                rg.UnMerge()
                rg.Value = text
                rg.Font.Name = "Segoe UI"
                rg.Font.Bold = True
                rg.Font.Size = 9
                rg.Font.Color = accent
                rg.Interior.Color = fill
                rg.HorizontalAlignment = -4108
                rg.VerticalAlignment = -4108
                for edge in (7, 8, 9, 10):
                    try:
                        b = rg.Borders(edge)
                        b.LineStyle = 1
                        b.Weight = 2
                        b.Color = accent
                    except Exception:
                        pass
            except Exception:
                pass

        def _style_input(merge_addr, top_left, wrap=False):
            try:
                rg = sh.Range(merge_addr)
                rg.UnMerge()
                rg.Merge()
            except Exception:
                pass
            try:
                rg = sh.Range(merge_addr)
                rg.Interior.Color = 0xFFFFFF
                rg.Font.Name = "Segoe UI"
                rg.Font.Size = 11
                rg.Font.Color = 0x1F1F1F
                rg.HorizontalAlignment = -4131
                rg.VerticalAlignment = -4160 if wrap else -4108
                rg.WrapText = wrap
                rg.IndentLevel = 1
                # card-like border
                for edge in (7, 8, 9, 10):
                    try:
                        b = rg.Borders(edge)
                        b.LineStyle = 1
                        b.Weight = 2
                        b.Color = 0xCFC6BA
                    except Exception:
                        pass
                tl = sh.Range(top_left)
                tl.Interior.Color = 0xFFFFFF
                tl.WrapText = wrap
            except Exception:
                pass

        # Canvas spacers
        try:
            sh.Range("B8:K8").Interior.Color = 0xF4F1EB
            sh.Range("B8:K8").Borders.LineStyle = -4142
            sh.Range("B8:K8").ClearContents()
        except Exception:
            pass

        # Sections (ordered, one per field). Fixed VBA cells: D5,D6,D7,D9
        _style_label("C5", "TO", 0xE8F1FB)
        _style_input("D5:K5", "D5", wrap=False)

        _style_label("C6", "CC", 0xEAF7F0)
        _style_input("D6:K6", "D6", wrap=False)

        _style_label("C7", "SUBJECT", 0xFFF4DF)
        _style_input("D7:K7", "D7", wrap=False)

        _style_label("C9", "BODY", 0xF3ECFF)
        _style_input("D9:K30", "D9", wrap=True)

        # Clear any stray/old content outside defined areas (keep header/instructions/fields)
        try:
            for rr in list(range(1, 2)) + list(range(10, 31)):
                # row 1 and below-body label column cleanup
                cell = sh.Range(f"C{rr}")
                if rr >= 10:
                    cell.ClearContents()
                    cell.Interior.Color = 0xF4F1EB
                    cell.Borders.LineStyle = -4142
            # clear legacy cells/spacers
            for addr in ("D8", "E8", "F8", "G8", "H8", "I8", "J8", "K8", "C8"):
                try:
                    sh.Range(addr).ClearContents()
                    sh.Range(addr).Interior.Color = 0xF4F1EB
                    sh.Range(addr).Borders.LineStyle = -4142
                except Exception:
                    pass
        except Exception:
            pass

        # Set values into the VBA-read cells
        try:
            sh.Range("D5").Value = field_values.get("to", "")
            sh.Range("D6").Value = field_values.get("cc", "")
            sh.Range("D7").Value = field_values.get("subject", "")
            sh.Range("D9").Value = field_values.get("body", "")
            sh.Range("D8").ClearContents()
        except Exception:
            pass

        # Window/view cleanup
        try:
            aw = wb.Application.ActiveWindow
            if aw is not None:
                aw.DisplayGridlines = False
                aw.DisplayHeadings = False
                aw.Zoom = 120
        except Exception:
            pass
        try:
            sh.Activate()
            sh.Range("D5").Select()
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
        accent=0x79411F,       # deep blue (BGR)
        accent_light=0xF3E6D7,
        badge_text="MANAGER EMAIL",
    )

    # Ensure hidden Director Email Template sheet exists with TO/CC/SUBJECT/BODY config.
    director_ws = get_or_create_sheet("Director Email Template")
    try:
        d_to = str(director_ws.Range("D5").Value or "").strip() if str(director_ws.Range("C5").Value or "").upper().startswith("TO") else str(director_ws.Range("D5").Value or "").strip()
        d_cc = str(director_ws.Range("D6").Value or "").strip()
        d_subj = str(director_ws.Range("D7").Value or "").strip()
        if d_subj == "" or d_subj == "Pending Pickup Report - Daily Dashboard":
            d_subj = "Pending Pickup Report - Monthly Dashboard"
        d_body = str(director_ws.Range("D9").Value or director_ws.Range("D8").Value or "").strip()
        if d_body == "" or d_body == "Please find the Pending Pickup daily dashboard below.":
            d_body = "Please find the Pending Pickup monthly dashboard below."
    except Exception:
        d_to, d_cc, d_subj, d_body = "", "", "Pending Pickup Report - Monthly Dashboard", "Please find the Pending Pickup monthly dashboard below."
    format_email_template_sheet(
        director_ws,
        "Director Email Template",
        {"to": d_to, "cc": d_cc, "subject": d_subj, "body": d_body},
        hide=True,
        accent=0x3E6B35,       # deep green (BGR)
        accent_light=0xE7F2E3,
        badge_text="DIRECTOR EMAIL",
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
