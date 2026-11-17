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
        # NOTE: VBA reads D5 (To), D6 (CC), D7 (Subject), D9 (Body) - keep those cells.
        try:
            if hide:
                sh.Visible = 0  # xlSheetHidden
        except Exception:
            pass

        # Reset sheet for a clean canvas
        try:
            for s in list(sh.Shapes):
                try:
                    s.Delete()
                except Exception:
                    pass
        except Exception:
            pass
        try:
            sh.Cells.Clear
            sh.Cells.UnMerge()
            sh.Cells.Interior.Color = 0xFAF7F2
            sh.Cells.Font.Name = "Segoe UI"
            sh.Cells.Font.Color = 0x222222
            sh.Cells.Borders.LineStyle = -4142  # xlLineStyleNone
            sh.Cells.WrapText = False
        except Exception:
            pass

        # Layout metrics (VBA reads D5=To, D6=CC, D7=Subject, D9=Body)
        try:
            sh.Columns("A").ColumnWidth = 1.0
            sh.Columns("B").ColumnWidth = 1.0
            sh.Columns("C").ColumnWidth = 11
            for col in ["D", "E", "F", "G", "H", "I", "J", "K"]:
                sh.Columns(col).ColumnWidth = 18
            sh.Rows("1").RowHeight = 10
            sh.Rows("2").RowHeight = 34
            sh.Rows("3").RowHeight = 18
            sh.Rows("4").RowHeight = 16   # spacer above TO
            sh.Rows("5").RowHeight = 44   # TO
            sh.Rows("6").RowHeight = 44   # CC
            sh.Rows("7").RowHeight = 44   # SUBJECT
            sh.Rows("8").RowHeight = 18   # spacer above BODY
            sh.Rows("9").RowHeight = 30   # BODY top
            for rr in range(10, 30):
                sh.Rows(str(rr)).RowHeight = 20
        except Exception:
            pass

        # Hide everything outside the design/input area
        try:
            sh.Rows("30:1048576").Hidden = True
            sh.Columns("L:XFD").Hidden = True
            sh.Columns("A").Hidden = True
            sh.Columns("B").Hidden = True
        except Exception:
            pass

        # Header card only (no giant board over all sections)
        try:
            header = sh.Shapes.AddShape(5, sh.Range("C2").Left - 2, sh.Range("C2").Top - 2,
                                        (sh.Range("K3").Left + sh.Range("K3").Width) - sh.Range("C2").Left + 4,
                                        (sh.Range("C3").Top + sh.Range("C3").Height) - sh.Range("C2").Top + 2)
            header.Name = "emailHeader"
            header.Fill.ForeColor.RGB = accent
            header.Line.Visible = 0
            header.Placement = 2
            try:
                header.Shadow.Type = 21
                header.Shadow.ForeColor.RGB = 0xC8C0B4
                header.Shadow.Transparency = 0.55
            except Exception:
                pass
        except Exception:
            pass

        try:
            badge = sh.Shapes.AddShape(5, sh.Range("I2").Left, sh.Range("C2").Top + 6,
                                       (sh.Range("K2").Left + sh.Range("K2").Width) - sh.Range("I2").Left - 2,
                                       22)
            badge.Name = "emailBadge"
            badge.Fill.ForeColor.RGB = accent_light
            badge.Line.Visible = 0
            badge.Placement = 2
            badge.TextFrame2.TextRange.Text = badge_text
            badge.TextFrame2.TextRange.Font.Name = "Segoe UI"
            badge.TextFrame2.TextRange.Font.Bold = True
            badge.TextFrame2.TextRange.Font.Size = 9
            badge.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = accent
            badge.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
            badge.TextFrame2.VerticalAnchor = 3
        except Exception:
            pass

        # Header text
        try:
            sh.Range("C2:K2").Merge()
            sh.Range("C3:K3").Merge()
            sh.Range("C2").Value = title
            sh.Range("C3").Value = "Enter details in each box below. Dashboard content is inserted automatically when you generate an email."
            sh.Range("C2").Font.Name = "Segoe UI"
            sh.Range("C2").Font.Bold = True
            sh.Range("C2").Font.Size = 18
            sh.Range("C2").Font.Color = 0xFFFFFF
            sh.Range("C2").HorizontalAlignment = -4131
            sh.Range("C2").VerticalAlignment = -4108
            sh.Range("C2").IndentLevel = 1
            sh.Range("C3").Font.Name = "Segoe UI"
            sh.Range("C3").Font.Size = 10
            sh.Range("C3").Font.Color = 0xF0E6D8
            sh.Range("C3").HorizontalAlignment = -4131
            sh.Range("C3").VerticalAlignment = -4108
            sh.Range("C3").IndentLevel = 1
        except Exception:
            pass

        # Individual separated sections: side label chip + isolated input box
        # Inset boxes vertically so sheet background shows between rows (sections look separate)
        for sec in [
            {"label": "TO", "row": 5, "rng": "D5:K5", "chip_fill": 0xE8F1FB},
            {"label": "CC", "row": 6, "rng": "D6:K6", "chip_fill": 0xEAF7F0},
            {"label": "SUBJECT", "row": 7, "rng": "D7:K7", "chip_fill": 0xFFF4DF},
        ]:
            row = sec["row"]
            label = sec["label"]
            rng_addr = sec["rng"]
            try:
                lrg = sh.Range(f"C{row}")
                chip = sh.Shapes.AddShape(5, lrg.Left, lrg.Top + 8, lrg.Width - 6, lrg.Height - 16)
                chip.Name = f"chip_{label}"
                chip.Fill.ForeColor.RGB = sec["chip_fill"]
                chip.Line.ForeColor.RGB = accent
                chip.Line.Weight = 0.9
                chip.Placement = 2
                chip.TextFrame2.TextRange.Text = label
                chip.TextFrame2.TextRange.Font.Name = "Segoe UI"
                chip.TextFrame2.TextRange.Font.Bold = True
                chip.TextFrame2.TextRange.Font.Size = 9
                chip.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = accent
                chip.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
                chip.TextFrame2.VerticalAnchor = 3
            except Exception:
                pass
            try:
                irg = sh.Range(rng_addr)
                # inset vertically so canvas shows between sections (individual boxes)
                box = sh.Shapes.AddShape(5, irg.Left, irg.Top + 6, irg.Width, irg.Height - 12)
                box.Name = f"input_{label}"
                box.Fill.ForeColor.RGB = 0xFFFFFF
                box.Line.ForeColor.RGB = accent
                box.Line.Weight = 1.5
                box.Placement = 2
                try:
                    box.Shadow.Type = 21
                    box.Shadow.ForeColor.RGB = 0xDDD6CC
                    box.Shadow.Transparency = 0.65
                except Exception:
                    pass
            except Exception:
                pass
            try:
                sh.Range(f"C{row}").Value = ""
            except Exception:
                pass

        # BODY: side label + large isolated input box
        try:
            body_chip = sh.Shapes.AddShape(5, sh.Range("C9").Left, sh.Range("C9").Top + 4,
                                           sh.Range("C9").Width - 6, 26)
            body_chip.Name = "chip_BODY"
            body_chip.Fill.ForeColor.RGB = 0xF3ECFF
            body_chip.Line.ForeColor.RGB = accent
            body_chip.Line.Weight = 0.9
            body_chip.Placement = 2
            body_chip.TextFrame2.TextRange.Text = "BODY"
            body_chip.TextFrame2.TextRange.Font.Name = "Segoe UI"
            body_chip.TextFrame2.TextRange.Font.Bold = True
            body_chip.TextFrame2.TextRange.Font.Size = 9
            body_chip.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = accent
            body_chip.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
            body_chip.TextFrame2.VerticalAnchor = 3
        except Exception:
            pass
        try:
            brg = sh.Range("D9:K29")
            brg.UnMerge()
            brg.Merge()
            body_box = sh.Shapes.AddShape(5, brg.Left, brg.Top, brg.Width, brg.Height)
            body_box.Name = "input_BODY"
            body_box.Fill.ForeColor.RGB = 0xFFFFFF
            body_box.Line.ForeColor.RGB = accent
            body_box.Line.Weight = 1.5
            body_box.Placement = 2
            try:
                body_box.Shadow.Type = 21
                body_box.Shadow.ForeColor.RGB = 0xDDD6CC
                body_box.Shadow.Transparency = 0.65
            except Exception:
                pass
        except Exception:
            pass

        # Input cell formatting (cells remain the actual editable fields, labels off to side only)
        try:
            for addr in ("D5", "D6", "D7"):
                rg = sh.Range(addr)
                rg.Interior.ColorIndex = 0  # transparent over shape
                rg.Font.Name = "Segoe UI"
                rg.Font.Size = 11
                rg.Font.Color = 0x1F1F1F
                rg.HorizontalAlignment = -4131
                rg.VerticalAlignment = -4108
                rg.WrapText = False
                rg.Borders.LineStyle = -4142
                rg.IndentLevel = 1
            brg = sh.Range("D9:K29")
            brg.Interior.ColorIndex = 0
            brg.Font.Name = "Segoe UI"
            brg.Font.Size = 11
            brg.Font.Color = 0x1F1F1F
            brg.HorizontalAlignment = -4131
            brg.VerticalAlignment = -4160  # xlTop
            brg.WrapText = True
            brg.Borders.LineStyle = -4142
            brg.IndentLevel = 1
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

        # Window/view cleanup
        try:
            aw = wb.Application.ActiveWindow
            if aw is not None:
                aw.DisplayGridlines = False
                aw.DisplayHeadings = False
                aw.Zoom = 125
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
        d_subj = str(director_ws.Range("D7").Value or "").strip() or "Pending Pickup Report - Daily Dashboard"
        d_body = str(director_ws.Range("D9").Value or director_ws.Range("D8").Value or "").strip() or "Please find the Pending Pickup daily dashboard below."
    except Exception:
        d_to, d_cc, d_subj, d_body = "", "", "Pending Pickup Report - Daily Dashboard", "Please find the Pending Pickup daily dashboard below."
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
