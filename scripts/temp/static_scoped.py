import os
import re
from collections import defaultdict

root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
src_dir = os.path.join(root, "vba")

KEYWORDS = set(
    """
    and as attribute byref byval call case const declare dim do each else elseif empty end endif enum eqv erase
    error event exit false for friend function get global gosub goto if imp implements in is let like long loop
    me mod new next not nothing null on option optional or paramarray preserve print private property public
    raiseevent redim rem resume return select set static step stop sub then to true type typeof until wend
    while with withevents xor boolean byte currency date decimal double integer object single string variant
    """.split()
)

BUILTINS = set(
    """
    application activesheet activeworkbook activewindow thisworkbook workbooks workbook worksheet worksheets
    sheets range cells rows columns shapes shape collection dictionary createobject getobject rgb msgbox
    inputbox dir doevents environ freefile open close print write put get seek eof lof savesetting getsetting
    deletesetting filedatetime format now date time timer left right mid len trim ucase lcase instr instrrev
    replace split join strcomp strreverse val cstr cdbl clng cint cdate cbool isnumeric isdate isempty isnull
    isarray isobject iserror ismissing array lbound ubound choose iif switch asc chr error err folder file
    files subfolders textstream outlook mailitem inspector chart series font borders interior trendlines axis
    axes ticklabels datalabels dataformat line fill textframe2 textrange paragraphformat adjustments
    items count keys exists add remove delete specialcells usedrange currentregion find findnext findprevious
    unmerge merge copy paste values formula formulas numberformat columnwidth rowheight autofit entirerow
    entirecolumn resize offset parent name names visible activate select clearcontents clearformats
    horizontalalignment verticalalignment wraptext indentlevel readingorder shrinktofit addindent
    chartobjects chartobject plotby source hasdatalabels haselegend hascharttitle charttitle majorunit
    minimumscale maximumscale hastitle legend position text string value2 address row column previous next
    after before filename readonly updatelinks fileformat savechanges lookin lookat what searchorder
    searchdirection matchbyte matchcase searchformat sort sortfield sortfields sortorder header orientation
    xlascending xldescending lefttoright applying_1 top width height appearance style
    mkdir kill append output input binary random worksheetfunction erl
    """.split()
)

decl_re = re.compile(r'^[ \t]*(?:(?:Public|Private|Global|Static|Dim|Const)\s+)+(.+)$', re.I | re.M)
proc_re = re.compile(
    r'^[ \t]*((?:(?:Public|Private|Friend|Static)\s+)?)(?:(?:Sub|Function|Property\s+(?:Get|Let|Set)))\s+([A-Za-z_]\w*)',
    re.I | re.M,
)
member_re = re.compile(r'\.\s*([A-Za-z_]\w*)')
named_arg_re = re.compile(r'([A-Za-z_]\w*)\s*:=')
id_re = re.compile(r'(?<![\w.])([A-Za-z_]\w*)\b')


def join_cont(text):
    return re.sub(r'[ \t]+_[ \t]*\r?\n', ' ', text)


def strip_code(text):
    text = join_cont(text)
    text = re.sub(r'^\s*Attribute\s+.*$', '', text, flags=re.M)
    text = re.sub(r'"(?:[^"]|"")*"', '""', text)
    text = re.sub(r"'(?:[^\n]*)", '', text)
    return text


def decl_entries(chunk):
    out = []
    for p in re.split(r',(?![^()]*\))', chunk):
        p = p.strip()
        if not p:
            continue
        p = re.sub(r'^(?:Withevents|Optional|ByVal|ByRef|ParamArray)\s+', '', p, flags=re.I)
        m = re.match(r'([A-Za-z_]\w*)', p)
        if m:
            out.append(m.group(1))
    return out


sources = {}
for f in os.listdir(src_dir):
    if f.endswith(".bas"):
        sources[f] = open(os.path.join(src_dir, f), encoding="utf-8", errors="replace").read()

module_locals = {}
module_public = {}
module_procs_public = set()
module_procs_all = defaultdict(set)

for f, text in sources.items():
    clean = strip_code(text)
    locals_, publics = set(), set()
    for m in decl_re.finditer(clean):
        line_start = clean.rfind("\n", 0, m.start()) + 1
        raw = clean[line_start:m.end()]
        is_public = bool(re.match(r'^[ \t]*(?:Public|Global)\b', raw, re.I))
        names = decl_entries(m.group(1))
        (publics if is_public else locals_).update(names)
    for m in proc_re.finditer(clean):
        vis = m.group(1).strip().lower()
        nm = m.group(2)
        module_procs_all[f].add(nm)
        if vis != "private":
            module_procs_public.add(nm)
    for pm in re.finditer(r'(?:Sub|Function)\s+[A-Za-z_]\w*\s*\(([^)]*)\)', clean, re.I):
        locals_.update(decl_entries(pm.group(1)))
    locals_.update(re.findall(r'^[ \t]*([A-Za-z_]\w*):', clean, re.M))
    module_locals[f] = locals_
    module_public[f] = publics

all_public = set()
for s in module_public.values():
    all_public |= s

print("Public declarations:", len(all_public))
for f, text in sources.items():
    clean = strip_code(text)
    masked = member_re.sub('.', clean)
    masked = named_arg_re.sub('', masked)
    visible = module_locals[f] | all_public | module_procs_public | module_procs_all[f]
    found = {}
    for m in id_re.finditer(masked):
        name = m.group(1)
        low = name.lower()
        if low in KEYWORDS or name in visible:
            continue
        if low in BUILTINS or low.startswith(("xl", "mso", "vb", "office")):
            continue
        line_no = masked[: m.start()].count("\n") + 1
        found.setdefault(name, []).append(line_no)
    if found:
        print(f"=== {f}")
        for name, lines in sorted(found.items()):
            print(f"  {name}: {lines[:12]}")
