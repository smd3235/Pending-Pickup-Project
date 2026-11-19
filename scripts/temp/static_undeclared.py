import os
import re

root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
src_dir = os.path.join(root, "vba")

KEYWORDS = set(
    """
    and as byref byval call case const declare dim do each else elseif empty end endif enum eqv erase error event exit
    false for friend function get global gosub goto if imp implements in is let like long loop me mod new next not
    nothing null on option optional or paramarray preserve print private property public raiseevent redim rem resume
    return select set static step stop sub then to true type typeof until wend while with withevents xor
    boolean byte currency date decimal double integer object single string variant
    attribute vb_name vb_crlf vbcrlf
    """.split()
)


def strip_code(text):
    text = re.sub(r'^\s*Attribute\s+.*$', '', text, flags=re.M)
    text = re.sub(r'"(?:[^"]|"")*"', '""', text)
    text = re.sub(r"'(?:[^\n]*)", '', text)
    return text


decl_re = re.compile(
    r'^[ \t]*(?:Dim|Private|Public|Global|Static|Const)[ \t]+([^\n]*)', re.I | re.M
)
proc_re = re.compile(
    r'^[ \t]*(?:(?:Public|Private|Friend|Static)[ \t]+)?(?:Sub|Function|Property[ \t]+(?:Get|Let|Set))[ \t]+([A-Za-z_]\w*)',
    re.I | re.M,
)
ident_re = re.compile(r'(?<![\w.])([A-Za-z_]\w*)\b(?=\s*[(=,)\n]|\s*$|\s*:)')


def split_names(chunk):
    parts = re.split(r',(?![^()]*\))', chunk)
    out = []
    for p in parts:
        p = p.split(" As ")[0].split(" as ")[0]
        m = re.match(r'\s*([A-Za-z_]\w*)\s*(\[?\(|\s*=)', p)
        if m:
            out.append(m.group(1))
            continue
        m = re.match(r'\s*([A-Za-z_]\w*)\s*$', p)
        if m:
            out.append(m.group(1))
    return out


declared = {}
procs = set()
module_of = {}
sources = {}
for f in os.listdir(src_dir):
    if not f.endswith(".bas"):
        continue
    text = open(os.path.join(src_dir, f), encoding="utf-8", errors="replace").read()
    sources[f] = text
    clean = strip_code(text)
    names = set()
    for m in decl_re.finditer(clean):
        for n in split_names(m.group(1)):
            names.add(n)
    for m in proc_re.finditer(clean):
        names.add(m.group(1))
        procs.add(m.group(1))
    # params
    for pm in re.finditer(r'(?:Sub|Function)[ \t]+[A-Za-z_]\w*[ \t]*\(([^)]*(?:\n[^)]*)*)\)', clean, re.I):
        for p in split_names(pm.group(1)):
            names.add(p)
    # With blocks: "With expr" no new names; For Each var declared already
    declared[f] = names
    module_of[f] = names

all_declared = set()
for s in declared.values():
    all_declared |= s

whitelist_prefixes = ("xl", "mso", "vb", "office")
whitelist_exact = {
    "Application", "ActiveSheet", "ActiveWorkbook", "ThisWorkbook", "Workbooks", "Workbook", "Worksheet",
    "Worksheets", "Sheets", "Range", "Cells", "Rows", "Columns", "Shapes", "Shape", "Collection",
    "Dictionary", "Scripting", "CreateObject", "GetObject", "RGB", "MsgBox", "InputBox", "Dir", "DoEvents",
    "Environ", "FreeFile", "Open", "Close", "Print", "Write", "Put", "Get", "Seek", "EOF", "LOF",
    "SaveSetting", "GetSetting", "DeleteSetting", "FileDateTime", "Format", "Now", "Date", "Time",
    "Timer", "Left", "Right", "Mid", "Len", "Trim", "UCase", "LCase", "InStr", "InStrRev", "Replace",
    "Split", "Join", "StrComp", "StrReverse", "Val", "CStr", "CDbl", "CLng", "CInt", "CDate", "CBool",
    "IsNumeric", "IsDate", "IsEmpty", "IsNull", "IsArray", "IsObject", "IsError", "IsMissing",
    "Array", "LBound", "UBound", "Choose", "IIf", "Switch", "Asc", "Chr", "Error", "Err",
    "Folder", "File", "Files", "SubFolders", "FSO", "TextStream", "Word", "MailItem", "Inspector",
    "Outlook", "Chart", "Series", "Font", "Borders", "Interior", "Trendlines", "Axis", "Axes",
    "TickLabels", "DataLabels", "Format", "Line", "Fill", "TextFrame2", "TextRange", "ParagraphFormat",
    "Borders", "Borders", "Sudoku",
}

unknown = {}
for f, text in sources.items():
    clean = strip_code(text)
    used = {}
    for m in ident_re.finditer(clean):
        name = m.group(1)
        low = name.lower()
        if name in KEYWORDS or low in KEYWORDS:
            continue
        if name in all_declared or name in procs or name in whitelist_exact:
            continue
        if low.startswith(whitelist_prefixes) or low.startswith("www"):
            continue
        line_no = clean[: m.start()].count("\n") + 1
        used.setdefault(name, []).append(line_no)
    unknown[f] = used

for f, used in unknown.items():
    if used:
        print(f"=== {f}")
        for name, lines in sorted(used.items()):
            print(f"  {name}: lines {lines[:10]}")
