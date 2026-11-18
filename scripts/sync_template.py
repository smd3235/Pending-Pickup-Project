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
parser.add_argument("--bas-dir", default=None, help="Directory containing .bas modules (default: vba/)")
parser.add_argument("--bas", action="append", default=[], help="Specific .bas file to sync (repeatable)")
parser.add_argument("--button", default="btnCreatePendingPickupReport")
parser.add_argument("--skip-package", action="store_true", help="Skip rebuilding Pending Pickup Template (no-macros).xlsx and Project-Local-Build.zip")
args = parser.parse_args()

script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)

template = args.template or os.path.join(project_root, "Pending Pickup Template.xlsm")
bas_dir = args.bas_dir or os.path.join(project_root, "vba")
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

    # ------------------------------------------------------------------
    # START HERE layout: hero, CTA card, email setup, how it works, good to know
    # ------------------------------------------------------------------
    ws = wb.Worksheets("START HERE")

    ACCENT = 0x1F4E79      # deep blue (BGR)
    ACCENT_DARK = 0x163A52 # darker blue for borders
    ACCENT_2 = 0x2E7D32    # deep green (BGR)
    CANVAS = 0xF4F7FB      # light blue-gray canvas
    CARD = 0xFFFFFF        # white cards
    CARD_BORDER = 0xC9D6E2 # subtle border
    TEXT = 0x1A1A2E        # near-black text
    MUTED = 0x5B6770       # muted gray
    LIGHT = 0xEAF0F6       # light panel background

    try:
        ws.Activate()
    except Exception:
        pass

    # Column/row scaffold — B:E is the content area (~104 width units)
    try:
        ws.Columns("A").ColumnWidth = 1.2
        ws.Columns("B").ColumnWidth = 26.0
        ws.Columns("C").ColumnWidth = 26.0
        ws.Columns("D").ColumnWidth = 26.0
        ws.Columns("E").ColumnWidth = 26.0
        ws.Columns("F").ColumnWidth = 12.0
        ws.Columns("G").ColumnWidth = 2.0
        ws.Rows("1").RowHeight = 4
        ws.Rows("2").RowHeight = 46    # hero title band
        ws.Rows("3").RowHeight = 20    # hero subtitle
        ws.Rows("4").RowHeight = 10
        ws.Rows("5").RowHeight = 18    # subtitle
        ws.Rows("6").RowHeight = 14
        ws.Rows("7").RowHeight = 20    # section label
        ws.Rows("8").RowHeight = 78    # CTA card (description)
        ws.Rows("9").RowHeight = 10
        ws.Rows("10").RowHeight = 56   # primary button row
        ws.Rows("11").RowHeight = 22
        ws.Rows("12").RowHeight = 20   # section label
        ws.Rows("13").RowHeight = 70   # email card
        ws.Rows("14").RowHeight = 10
        ws.Rows("15").RowHeight = 42   # email buttons row
        ws.Rows("16").RowHeight = 22
        ws.Rows("17").RowHeight = 20   # section label
        ws.Rows("18").RowHeight = 18   # step 1
        ws.Rows("19").RowHeight = 18   # step 2
        ws.Rows("20").RowHeight = 18   # step 3
        ws.Rows("21").RowHeight = 18   # step 4
        ws.Rows("22").RowHeight = 18   # step 5
        ws.Rows("23").RowHeight = 14
        ws.Rows("24").RowHeight = 20   # section label
        ws.Rows("25").RowHeight = 18   # note 1
        ws.Rows("26").RowHeight = 18   # note 2
        ws.Rows("27").RowHeight = 18   # note 3
        ws.Rows("28").RowHeight = 18   # note 4
        ws.Rows("29").RowHeight = 14
        ws.Rows("30").RowHeight = 18   # footer
    except Exception:
        pass

    # Reset visible content/format region for a clean canvas
    try:
        rng = ws.Range("B1:E30")
        rng.ClearContents()
        rng.UnMerge()
        rng.Interior.ColorIndex = -4142  # xlColorIndexNone
        rng.Borders.LineStyle = -4142
        ws.Cells.Font.Name = "Segoe UI"
    except Exception:
        pass

    try:
        ws.Range("A1:G30").Interior.Color = CANVAS
    except Exception:
        pass

    def _merge(addr, value=None):
        try:
            rg = ws.Range(addr)
            rg.UnMerge()
            rg.Merge()
            if value is not None:
                rg.Value = value
            return rg
        except Exception:
            try:
                return ws.Range(addr.split(":")[0])
            except Exception:
                return None

    def _set_font(addr, size=11, bold=False, color=TEXT, name="Segoe UI"):
        try:
            rg = ws.Range(addr)
            rg.Font.Name = name
            rg.Font.Size = size
            rg.Font.Bold = bold
            rg.Font.Color = color
        except Exception:
            pass

    def _box(addr, fill=CARD, border=CARD_BORDER, weight=1):
        try:
            rg = ws.Range(addr)
            rg.Interior.Color = fill
            for edge in (7, 8, 9, 10):
                b = rg.Borders(edge)
                b.LineStyle = 1
                b.Weight = weight
                b.Color = border
        except Exception:
            pass

    # ---------- HERO BAND ----------
    try:
        _merge("B2:E3", "Pending Pickup Report Builder")
        ws.Range("B2:E3").Interior.Color = ACCENT
        for edge in (7, 8, 9, 10):
            b = ws.Range("B2:E3").Borders(edge)
            b.LineStyle = 1
            b.Weight = 2
            b.Color = ACCENT_DARK
        _set_font("B2:E2", size=26, bold=True, color=0xFFFFFF, name="Segoe UI Semibold")
        ws.Range("B2").HorizontalAlignment = -4131  # center
        ws.Range("B2").VerticalAlignment = -4108    # middle
        ws.Range("B2").IndentLevel = 1
    except Exception:
        pass

    try:
        _merge("B3:E3", "Generate the monthly Pending Pickup workbook and prepare manager / director email drafts.")
        _set_font("B3:E3", size=11, bold=False, color=0xDCE8F2, name="Segoe UI")
        ws.Range("B3").HorizontalAlignment = -4131
        ws.Range("B3").VerticalAlignment = -4108
        ws.Range("B3").IndentLevel = 1
    except Exception:
        pass

    # ---------- SUBTITLE ----------
    try:
        _merge("B5:E5", "Start here to build a new Pending Pickup report from your current export.")
        _set_font("B5:E5", size=11, bold=False, color=MUTED)
        ws.Range("B5").HorizontalAlignment = -4131
        ws.Range("B5").IndentLevel = 1
    except Exception:
        pass

    # ---------- SECTION: CREATE REPORT ----------
    try:
        _merge("B7:E7", "STEP 1  •  CREATE REPORT")
        _set_font("B7:E7", size=10, bold=True, color=MUTED)
        ws.Range("B7").IndentLevel = 1
    except Exception:
        pass

    # CTA card — description
    try:
        _box("B8:E9", fill=CARD, border=CARD_BORDER, weight=1)
        _merge("B8:E9", "Create a new Pending Pickup report from the current All Open Treatment Plans export." + "\n" + "You will choose the source file, optionally select last month\u2019s report for comparison, and pick a save location.")
        ws.Range("B8:E9").Font.Size = 10.5
        ws.Range("B8:E9").Font.Color = TEXT
        ws.Range("B8:E9").HorizontalAlignment = -4131
        ws.Range("B8:E9").VerticalAlignment = -4108
        ws.Range("B8:E9").WrapText = True
        ws.Range("B8:E9").IndentLevel = 1
    except Exception:
        pass

    # ---------- PRIMARY CTA BUTTON ----------
    try:
        anchor = ws.Range("B8")
        _make_button(
            button_name, "CreatePendingPickupReport",
            anchor.Left + 14, anchor.Top + 86,
            340, 52,
            "CREATE PENDING PICKUP REPORT",
            ACCENT, ACCENT_DARK, font_size=14, bold=True,
        )
    except Exception:
        pass

    # ---------- SECTION: EMAIL SETUP ----------
    try:
        _merge("B12:E12", "STEP 2  •  EMAIL TEMPLATES")
        _set_font("B12:E12", size=10, bold=True, color=MUTED)
        ws.Range("B12").IndentLevel = 1
    except Exception:
        pass

    # Email setup card — description
    try:
        _box("B13:E13", fill=CARD, border=CARD_BORDER, weight=1)
        _merge("B13:E13", "Configure the email templates used when creating drafts from generated reports." + "\n" + "Manager: per-manager tab emails  \u2022  Director: monthly dashboard email.")
        ws.Range("B13:E13").Font.Size = 10
        ws.Range("B13:E13").Font.Color = TEXT
        ws.Range("B13:E13").HorizontalAlignment = -4131
        ws.Range("B13:E13").VerticalAlignment = -4108
        ws.Range("B13:E13").WrapText = True
        ws.Range("B13:E13").IndentLevel = 1
    except Exception:
        pass

    # Two email setup buttons side by side
    try:
        card = ws.Range("B13")
        left = card.Left + 14
        top = card.Top + 70
        w = (card.Width - 42) / 2.0
        h = 34
        _make_button("btnSetupEmailCC", "SetupEmailCC", left, top, w, h, "Setup Manager Email", 0x5A86B8, 0x355A86, font_size=10)
        _make_button("btnSetupDirectorEmail", "SetupDirectorEmail", left + w + 14, top, w, h, "Setup Director Email", ACCENT_2, 0x1E5620, font_size=10)
    except Exception:
        pass

    # ---------- SECTION: HOW IT WORKS ----------
    try:
        _merge("B17:E17", "STEP 3  •  HOW IT WORKS")
        _set_font("B17:E17", size=10, bold=True, color=MUTED)
        ws.Range("B17").IndentLevel = 1
    except Exception:
        pass

    steps = [
        ("1", "Click CREATE PENDING PICKUP REPORT."),
        ("2", "Select the current All Open Treatment Plans export."),
        ("3", "Optionally select last month’s report for comparison."),
        ("4", "Choose a save location (reports are saved in Reports/ by default)."),
        ("5", "Use the Email buttons on generated tabs to create Outlook drafts."),
    ]
    try:
        row = 18
        for num, text in steps:
            addr_n = "D" + str(row)
            addr_t = "E" + str(row) + ":E" + str(row)
            ws.Range(addr_n).Value = num
            _merge(addr_t, text)
            ws.Range(addr_n).Font.Bold = True
            ws.Range(addr_n).Font.Color = ACCENT
            ws.Range(addr_n).HorizontalAlignment = -4108
            ws.Range(addr_n).VerticalAlignment = -4108
            ws.Range("E" + str(row)).Font.Size = 10
            ws.Range("E" + str(row)).Font.Color = TEXT
            ws.Range("E" + str(row)).VerticalAlignment = -4108
            ws.Range("E" + str(row)).IndentLevel = 0
            row += 1
        ws.Range("D23:E23").ClearContents()
    except Exception:
        pass

    # ---------- SECTION: GOOD TO KNOW ----------
    try:
        _merge("B24:E24", "GOOD TO KNOW")
        _set_font("B24:E24", size=10, bold=True, color=MUTED)
        ws.Range("B24").IndentLevel = 1
    except Exception:
        pass

    _box("B25:E28", fill=LIGHT, border=CARD_BORDER, weight=1)
    notes = [
        ("B25", "•  Only rows with Treatment Status ‘Pending Pickup’ are included."),
        ("B26", "•  Pickup Age is days since AUDIT: TP End Date — older means higher priority."),
        ("B27", "•  A missing AUDIT: TP End Date is shown as ‘MISSING DATE’."),
        ("B28", "•  Edit email defaults in hidden template tabs via Setup Email buttons."),
    ]
    for addr, val in notes:
        try:
            _merge(addr + ":E" + addr[1:], val)
            ws.Range(addr).Font.Size = 9.5
            ws.Range(addr).Font.Color = TEXT
            ws.Range(addr).VerticalAlignment = -4108
            ws.Range(addr).IndentLevel = 1
        except Exception:
            pass

    # ---------- FOOTER ----------
    try:
        _merge("B30:E30", "v1.0  •  Pending Pickup Report Builder")
        _set_font("B30:E30", size=9, bold=False, color=MUTED)
        ws.Range("B30").HorizontalAlignment = -4152  # right
        ws.Range("B30").IndentLevel = 1
    except Exception:
        pass

    # Remove/replace old START HERE shapes to avoid stale overlays
    try:
        for s in list(ws.Shapes):
            nm = str(getattr(s, "Name", ""))
            if nm in {button_name, "btnSetupEmailCC", "btnSetupDirectorEmail"} or nm.lower().startswith("btncreate") or nm.lower().startswith("btnsetup"):
                try:
                    s.Delete()
                except Exception:
                    pass
    except Exception:
        pass

    def _make_button(name, macro, left, top, width, height, text, fill, line, font_size=11, bold=True):
        b = ws.Shapes.AddShape(5, left, top, width, height)  # msoShapeRoundedRectangle
        b.Name = name
        b.OnAction = macro
        b.Fill.ForeColor.RGB = fill
        b.Line.ForeColor.RGB = line
        b.Line.Weight = 1.25
        try:
            b.TextFrame2.TextRange.Text = text
            b.TextFrame2.TextRange.Font.Name = "Segoe UI"
            b.TextFrame2.TextRange.Font.Bold = bold
            b.TextFrame2.TextRange.Font.Size = font_size
            b.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF
            b.TextFrame2.TextRange.ParagraphFormat.Alignment = 2  # center
            b.TextFrame2.VerticalAnchor = 3  # middle
        except Exception:
            pass
        try:
            b.Placement = 3  # free floating
        except Exception:
            pass
        return b
    # View cleanup
    try:
        aw = wb.Application.ActiveWindow
        if aw is not None:
            aw.DisplayGridlines = False
            aw.DisplayHeadings = False
            aw.Zoom = 115
        ws.Range("B2").Select()
    except Exception:
        pass

    def format_email_template_sheet(sh, title, field_values, hide=True, accent=0x79411F, accent_light=0xF2E8D5, badge_text="EMAIL TEMPLATE", auto_send="", attach_report=""):
        # field_values: dict with keys to, cc, subject, body
        # NOTE: VBA readers expect D5 (To), D7 (CC), D9 (Subject), D11 (Body) - keep these cells.
        # Additional email-behavior cells: D35 (Auto Send "1"=send), D36 (Attach Report "1"=attach).
        try:
            if hide:
                sh.Visible = 0  # xlSheetHidden
        except Exception:
            pass

        # Preserve existing email-behavior cell values before the full clear.
        existing_auto_send = ""
        existing_attach_report = ""
        try:
            existing_auto_send = str(sh.Range("D35").Value or "").strip()
        except Exception:
            pass
        try:
            existing_attach_report = str(sh.Range("D36").Value or "").strip()
        except Exception:
            pass
        if auto_send == "":
            auto_send = existing_auto_send
        if attach_report == "":
            attach_report = existing_attach_report

        # Derived palette
        def _blend(hex_color, white_mix=0.82):
            # simple lighten toward white in BGR int
            r = (hex_color >> 0) & 0xFF
            g = (hex_color >> 8) & 0xFF
            b = (hex_color >> 16) & 0xFF
            nr = int(r + (255 - r) * white_mix)
            ng = int(g + (255 - g) * white_mix)
            nb = int(b + (255 - b) * white_mix)
            return nr | (ng << 8) | (nb << 16)

        CANVAS = 0xF3F1EC
        PANEL = 0xFFFFFF
        PANEL_BORDER = 0xD7D1C6
        MUTED = 0x5A5348
        LABEL_TEXT = accent
        LABEL_BG = accent_light if accent_light else _blend(accent, 0.86)
        HINT_BG = _blend(accent, 0.93)

        # Reset sheet for a clean, shape-free canvas
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
            sh.Cells.Interior.Color = CANVAS
            sh.Cells.Font.Name = "Segoe UI"
            sh.Cells.Font.Size = 11
            sh.Cells.Font.Color = 0x1F1F1F
            sh.Cells.Borders.LineStyle = -4142  # xlLineStyleNone
            sh.Cells.WrapText = False
            sh.Cells.RowHeight = 18
        except Exception:
            pass

        # Layout (preserve VBA-read cells): header 2-3, badge/instructions 4, section TO 5, gap 6, CC 7, gap 8, SUBJECT 9, gap 10, BODY 11:31
        try:
            sh.Columns("A").ColumnWidth = 1.0
            sh.Columns("B").ColumnWidth = 1.6
            sh.Columns("C").ColumnWidth = 13.5
            for col in ["D", "E", "F", "G", "H", "I", "J", "K"]:
                sh.Columns(col).ColumnWidth = 19.5
            sh.Rows("1").RowHeight = 6
            sh.Rows("2").RowHeight = 40   # title
            sh.Rows("3").RowHeight = 22   # subtitle
            sh.Rows("4").RowHeight = 20   # badge + hints
            sh.Rows("5").RowHeight = 34   # TO (D5)
            sh.Rows("6").RowHeight = 8    # spacer
            sh.Rows("7").RowHeight = 34   # CC (D7)
            sh.Rows("8").RowHeight = 8    # spacer
            sh.Rows("9").RowHeight = 34   # SUBJECT (D9)
            sh.Rows("10").RowHeight = 12  # spacer before BODY
            sh.Rows("11").RowHeight = 28  # BODY label row top
            for rr in range(12, 33):
                sh.Rows(str(rr)).RowHeight = 19.5
            sh.Rows("33").RowHeight = 8
            sh.Rows("34").RowHeight = 34   # AUTO SEND (D35)
            sh.Rows("35").RowHeight = 34   # ATTACH REPORT (D36)
            sh.Rows("36").RowHeight = 8    # spacer
        except Exception:
            pass

        try:
            sh.Rows("37:1048576").Hidden = True
            sh.Columns("L:XFD").Hidden = True
            sh.Columns("A").Hidden = True
        except Exception:
            pass

        # Hero/header band
        try:
            sh.Range("B2:K2").Merge()
            sh.Range("B3:K3").Merge()
            sh.Range("B2:K3").Interior.Color = accent
            sh.Range("B2").Value = title
            sh.Range("B3").Value = "Configure recipients and message content. The report is pasted at {{REPORT}} when you send."
            sh.Range("B2").Font.Name = "Segoe UI Semibold"
            sh.Range("B2").Font.Bold = True
            sh.Range("B2").Font.Size = 20
            sh.Range("B2").Font.Color = 0xFFFFFF
            sh.Range("B2").HorizontalAlignment = -4131
            sh.Range("B2").VerticalAlignment = -4108
            sh.Range("B2").IndentLevel = 1
            sh.Range("B3").Font.Name = "Segoe UI"
            sh.Range("B3").Font.Size = 10
            sh.Range("B3").Font.Color = 0xF7F2EA
            sh.Range("B3").HorizontalAlignment = -4131
            sh.Range("B3").VerticalAlignment = -4108
            sh.Range("B3").IndentLevel = 1
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

        # Return to START HERE button (anchored in header right)
        try:
            btn_cell = sh.Range("J2:K2")
            # ensure clean header right area
            try:
                sh.Range("J2:K2").UnMerge()
                sh.Range("J2:K2").Interior.Color = accent
                sh.Range("J2:K2").ClearContents()
            except Exception:
                pass
            btn = sh.Shapes.AddShape(5, btn_cell.Left + 4, btn_cell.Top + 6, btn_cell.Width - 8, 26)
            btn.Name = "btnReturnStartHere"
            btn.OnAction = "ReturnToStartHere"
            btn.Fill.ForeColor.RGB = accent_light
            btn.Line.ForeColor.RGB = 0xFFFFFF
            btn.Line.Weight = 1.25
            btn.TextFrame2.TextRange.Text = "↩  Return to START HERE"
            btn.TextFrame2.TextRange.Font.Name = "Segoe UI"
            btn.TextFrame2.TextRange.Font.Bold = True
            btn.TextFrame2.TextRange.Font.Size = 9
            btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = accent
            btn.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
            btn.TextFrame2.VerticalAnchor = 3
            try:
                btn.Placement = 2
            except Exception:
                pass
        except Exception:
            pass

        # Info strip: badge + quick tips
        try:
            sh.Range("B4:C4").Merge()
            sh.Range("D4:K4").Merge()
            sh.Range("B4").Value = f"  {badge_text}  "
            sh.Range("B4").Font.Name = "Segoe UI Semibold"
            sh.Range("B4").Font.Bold = True
            sh.Range("B4").Font.Size = 9
            sh.Range("B4").Font.Color = accent
            sh.Range("B4").Interior.Color = HINT_BG
            sh.Range("B4").HorizontalAlignment = -4108
            sh.Range("B4").VerticalAlignment = -4108
            sh.Range("D4").Value = "TO/CC: use semicolons (;)  •  BODY: keep {{REPORT}} (text above = intro, below = closing)  •  New line: Alt+Enter"
            sh.Range("D4").Font.Name = "Segoe UI"
            sh.Range("D4").Font.Size = 9
            sh.Range("D4").Font.Italic = True
            sh.Range("D4").Font.Color = MUTED
            sh.Range("D4").Interior.Color = 0xEAE7DF
            sh.Range("D4").HorizontalAlignment = -4131
            sh.Range("D4").VerticalAlignment = -4108
            sh.Range("D4").IndentLevel = 1
            for rg_addr in ("B4:C4", "D4:K4"):
                for edge in (7, 8, 9, 10):
                    try:
                        b = sh.Range(rg_addr).Borders(edge)
                        b.LineStyle = 1
                        b.Weight = 2
                        b.Color = PANEL_BORDER
                    except Exception:
                        pass
        except Exception:
            pass

        def _style_field(row, label, input_addr, merge_addr, hint, label_fill, wrap=False):
            try:
                # label cell
                lr = sh.Range(f"C{row}")
                lr.UnMerge()
                lr.Value = label
                lr.Font.Name = "Segoe UI Semibold"
                lr.Font.Bold = True
                lr.Font.Size = 9
                lr.Font.Color = LABEL_TEXT
                lr.Interior.Color = label_fill
                lr.HorizontalAlignment = -4108
                lr.VerticalAlignment = -4108
                for edge in (7, 8, 9, 10):
                    b = lr.Borders(edge)
                    b.LineStyle = 1
                    b.Weight = 2
                    b.Color = accent
            except Exception:
                pass
            try:
                ir = sh.Range(merge_addr)
                ir.UnMerge()
                ir.Merge()
                ir.Interior.Color = PANEL
                ir.Font.Name = "Segoe UI"
                ir.Font.Size = 11
                ir.Font.Color = 0x1F1F1F
                ir.HorizontalAlignment = -4131
                ir.VerticalAlignment = -4160 if wrap else -4108
                ir.WrapText = wrap
                ir.IndentLevel = 1
                for edge in (7, 8, 9, 10):
                    b = ir.Borders(edge)
                    b.LineStyle = 1
                    b.Weight = 2
                    b.Color = PANEL_BORDER
                tl = sh.Range(input_addr)
                tl.Interior.Color = PANEL
                tl.WrapText = wrap
            except Exception:
                pass

        # Spacer rows
        for r in (6, 8, 10, 33):
            try:
                sh.Range(f"B{r}:K{r}").Interior.Color = CANVAS
                sh.Range(f"B{r}:K{r}").Borders.LineStyle = -4142
                sh.Range(f"B{r}:K{r}").ClearContents()
            except Exception:
                pass

        # Fields (fixed VBA cells)
        _style_field(5, "TO", "D5", "D5:K5", "", 0xE8F1FB, wrap=False)
        _style_field(7, "CC", "D7", "D7:K7", "", 0xEAF7F0, wrap=False)
        _style_field(9, "SUBJECT", "D9", "D9:K9", "", 0xFFF4DF, wrap=False)

        # Email-behavior toggles (read by VBA from registry EMAIL_REG_KEY)
        _style_field(34, "AUTO SEND", "D35", "D35:K35", 'Enter "1" to auto-send (leave blank to open as draft)', 0xFFF4DF, wrap=False)
        _style_field(35, "ATTACH REPORT", "D36", "D36:K36", 'Enter "1" to attach the generated report .xlsx', 0xE8F1FB, wrap=False)

        # BODY section: label block + large input area
        try:
            sh.Range("C11:C31").UnMerge()
            sh.Range("C11:C31").Merge()
            bl = sh.Range("C11")
            bl.Value = "BODY"
            bl.Font.Name = "Segoe UI Semibold"
            bl.Font.Bold = True
            bl.Font.Size = 9
            bl.Font.Color = LABEL_TEXT
            bl.Interior.Color = 0xF3ECFF
            bl.HorizontalAlignment = -4108
            bl.VerticalAlignment = -4160
            bl.Orientation = 0
            for edge in (7, 8, 9, 10):
                b = sh.Range("C11:C31").Borders(edge)
                b.LineStyle = 1
                b.Weight = 2
                b.Color = accent
        except Exception:
            pass
        try:
            br = sh.Range("D11:K31")
            br.UnMerge()
            br.Merge()
            br.Interior.Color = PANEL
            br.Font.Name = "Segoe UI"
            br.Font.Size = 11
            br.Font.Color = 0x1F1F1F
            br.HorizontalAlignment = -4131
            br.VerticalAlignment = -4160
            br.WrapText = True
            br.IndentLevel = 1
            for edge in (7, 8, 9, 10):
                b = br.Borders(edge)
                b.LineStyle = 1
                b.Weight = 2
                b.Color = PANEL_BORDER
            sh.Range("D11").WrapText = True
        except Exception:
            pass

        # Right-margin note card for body guidance
        # (kept inside canvas by using a subtle hint above body instead of extra columns)
        try:
            # light inner header tint for body label top
            sh.Range("C11").Interior.Color = 0xEBDDFF
        except Exception:
            pass

        # Clear stray/old content outside defined areas
        try:
            for addr in ("C6", "D6", "C8", "D8", "C10", "D10", "B5", "B7", "B9", "B11"):
                try:
                    sh.Range(addr).ClearContents()
                    sh.Range(addr).Interior.Color = CANVAS
                    sh.Range(addr).Borders.LineStyle = -4142
                except Exception:
                    pass
        except Exception:
            pass

        # Set values into the VBA-read cells
        try:
            sh.Range("D5").Value = field_values.get("to", "")
            sh.Range("D7").Value = field_values.get("cc", "")
            sh.Range("D9").Value = field_values.get("subject", "")
            sh.Range("D11").Value = field_values.get("body", "")
            sh.Range("D35").Value = auto_send
            sh.Range("D36").Value = attach_report
            for addr in ("D6", "D8", "D10"):
                try:
                    sh.Range(addr).ClearContents()
                except Exception:
                    pass
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
        c5 = str(mgr_ws.Range("C5").Value or "").upper()
        c6 = str(mgr_ws.Range("C6").Value or "").upper()
        c7 = str(mgr_ws.Range("C7").Value or "").upper()
        c9 = str(mgr_ws.Range("C9").Value or "").upper()

        # New spaced layout (D5,D7,D9,D11)
        if c7.startswith("CC") and c9.startswith("SUBJ"):
            mgr_to = str(mgr_ws.Range("D5").Value or "").strip()
            mgr_cc = str(mgr_ws.Range("D7").Value or "").strip()
            mgr_subj = str(mgr_ws.Range("D9").Value or "").strip() or "Pending Pickup Report - Action Required"
            mgr_body = str(mgr_ws.Range("D11").Value or mgr_ws.Range("D9").Value or "").strip()
        # Recent layout (D5,D6,D7,D9)
        elif c6.startswith("CC") and c7.startswith("SUBJ"):
            mgr_to = str(mgr_ws.Range("D5").Value or "").strip()
            mgr_cc = str(mgr_ws.Range("D6").Value or "").strip()
            mgr_subj = str(mgr_ws.Range("D7").Value or "").strip() or "Pending Pickup Report - Action Required"
            mgr_body = str(mgr_ws.Range("D9").Value or "").strip()
        # Old layout (CC/Subject/Body at D5,D6,D7)
        elif c5.startswith("CC"):
            mgr_to = ""
            mgr_cc = str(mgr_ws.Range("D5").Value or "").strip()
            mgr_subj = str(mgr_ws.Range("D6").Value or "").strip() or "Pending Pickup Report - Action Required"
            mgr_body = str(mgr_ws.Range("D7").Value or "").strip()
        else:
            mgr_to = str(mgr_ws.Range("D5").Value or "").strip()
            mgr_cc = str(mgr_ws.Range("D7").Value or mgr_ws.Range("D6").Value or "").strip()
            mgr_subj = str(mgr_ws.Range("D9").Value or mgr_ws.Range("D7").Value or "").strip() or "Pending Pickup Report - Action Required"
            mgr_body = str(mgr_ws.Range("D11").Value or mgr_ws.Range("D9").Value or "").strip()

        if not mgr_body or mgr_body == "Please review the Pending Pickup report snapshot below.":
            mgr_body = "Please review the Pending Pickup report below.\n\n{{REPORT}}\n\nPlease review and respond as needed.\nPending Pickup Report System"
    except Exception:
        mgr_to, mgr_cc, mgr_subj, mgr_body = "", "", "Pending Pickup Report - Action Required", "Please review the Pending Pickup report below.\n\n{{REPORT}}\n\nPlease review and respond as needed.\nPending Pickup Report System"
    try:
        mgr_auto_send = str(mgr_ws.Range("D35").Value or "").strip()
        mgr_attach_report = str(mgr_ws.Range("D36").Value or "").strip()
    except Exception:
        mgr_auto_send, mgr_attach_report = "", ""
    format_email_template_sheet(
        mgr_ws,
        "Manager Email Template",
        {"to": mgr_to, "cc": mgr_cc, "subject": mgr_subj, "body": mgr_body},
        hide=True,
        accent=0x79411F,       # deep blue (BGR)
        accent_light=0xF3E6D7,
        badge_text="MANAGER EMAIL",
        auto_send=mgr_auto_send,
        attach_report=mgr_attach_report,
    )

    # Ensure hidden Director Email Template sheet exists with TO/CC/SUBJECT/BODY config.
    director_ws = get_or_create_sheet("Director Email Template")
    try:
        c5 = str(director_ws.Range("C5").Value or "").upper()
        c6 = str(director_ws.Range("C6").Value or "").upper()
        c7 = str(director_ws.Range("C7").Value or "").upper()
        c9 = str(director_ws.Range("C9").Value or "").upper()

        # New spaced layout (D5,D7,D9,D11)
        if c7.startswith("CC") and c9.startswith("SUBJ"):
            d_to = str(director_ws.Range("D5").Value or "").strip()
            d_cc = str(director_ws.Range("D7").Value or "").strip()
            d_subj = str(director_ws.Range("D9").Value or "").strip()
            d_body = str(director_ws.Range("D11").Value or director_ws.Range("D9").Value or "").strip()
        # Recent layout (D5,D6,D7,D9)
        elif c6.startswith("CC") and c7.startswith("SUBJ"):
            d_to = str(director_ws.Range("D5").Value or "").strip()
            d_cc = str(director_ws.Range("D6").Value or "").strip()
            d_subj = str(director_ws.Range("D7").Value or "").strip()
            d_body = str(director_ws.Range("D9").Value or "").strip()
        else:
            d_to = str(director_ws.Range("D5").Value or "").strip()
            d_cc = str(director_ws.Range("D7").Value or director_ws.Range("D6").Value or "").strip()
            d_subj = str(director_ws.Range("D9").Value or director_ws.Range("D7").Value or "").strip()
            d_body = str(director_ws.Range("D11").Value or director_ws.Range("D9").Value or "").strip()

        if d_subj == "" or d_subj == "Pending Pickup Report - Daily Dashboard":
            d_subj = "Pending Pickup Report - Monthly Dashboard"
        if d_body == "" or d_body == "Please find the Pending Pickup daily dashboard below." or d_body == "Please find the Pending Pickup monthly dashboard below.":
            d_body = "Please find the Pending Pickup monthly dashboard below.\n\n{{REPORT}}\n\nPending Pickup Report System"
    except Exception:
        d_to, d_cc, d_subj, d_body = "", "", "Pending Pickup Report - Monthly Dashboard", "Please find the Pending Pickup monthly dashboard below.\n\n{{REPORT}}\n\nPending Pickup Report System"
    try:
        d_auto_send = str(director_ws.Range("D35").Value or "").strip()
        d_attach_report = str(director_ws.Range("D36").Value or "").strip()
    except Exception:
        d_auto_send, d_attach_report = "", ""
    format_email_template_sheet(
        director_ws,
        "Director Email Template",
        {"to": d_to, "cc": d_cc, "subject": d_subj, "body": d_body},
        hide=True,
        accent=0x3E6B35,       # deep green (BGR)
        accent_light=0xE7F2E3,
        badge_text="DIRECTOR EMAIL",
        auto_send=d_auto_send,
        attach_report=d_attach_report,
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

    # Collect all VBA source modules in vba/ (.bas, .cls, .frm)
    module_files = []
    for ext in ("*.bas", "*.cls", "*.frm"):
        module_files.extend(glob.glob(os.path.join(project_root, "vba", ext)))
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

    # Build a separate zip of the macro-enabled template (.xlsm). The .bas
    # modules are already embedded in this file by the sync above, so this
    # zip carries only the .xlsm and is created in tandem with the local
    # build package above.
    macro_zip_path = os.path.join(project_root, "Pending Pickup Template (macro).zip")
    if os.path.exists(macro_zip_path):
        os.remove(macro_zip_path)
    with zipfile.ZipFile(macro_zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.write(template, arcname=os.path.basename(template))
    print(f"Built macro template package: {macro_zip_path}")
    print("  Includes:")
    print(f"    - {os.path.basename(template)}")
