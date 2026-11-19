import ctypes
import ctypes.wintypes as wt
import os
import threading
import time

import psutil
import pythoncom
import win32com.client

user32 = ctypes.windll.user32
EnumWindows = user32.EnumWindows
EnumChildWindows = user32.EnumChildWindows
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
GetWindowTextW = user32.GetWindowTextW
GetClassNameW = user32.GetClassNameW
IsWindowVisible = user32.IsWindowVisible

EnumProc = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)

root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
temp = os.path.join(root, "scripts", "temp")

VARS = """Public mAuditStartTime As Date
Public mAuditSourcePath As String
Public mAuditOutputPath As String
Public mAuditTotalRecords As Long
Public mAuditPendingRecords As Long
Public mAuditNewRecords As Long
Public mAuditClearedRecords As Long
Public mAuditStillRecords As Long
Public mAuditMissingDates As Long
Public mAuditDuplicateKeys As Long
Public mAuditMissingManagers As Long
Public mAuditMissingStatus As Long
Public mAuditFutureDates As Long
Public mAuditVeryOldDates As Long
Public mAuditErrorCount As Long
Public mAuditIssues As Collection
Public mRunLogPath As String
Public mInReportRun As Boolean
Public mRepSeenDict As Object"""

dialogs = []


def win_text(hwnd):
    n = user32.GetWindowTextLengthW(hwnd)
    b = ctypes.create_unicode_buffer(n + 1)
    user32.GetWindowTextW(hwnd, b, n + 1)
    return b.value


def win_class(hwnd):
    b = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, b, 256)
    return b.value


def kill_excel():
    for p in psutil.process_iter(["name"]):
        if p.info["name"] and p.info["name"].lower() == "excel.exe":
            p.kill()


def watch(pid, stop):
    while not stop.is_set():
        def cb(hwnd, _):
            p = wt.DWORD()
            GetWindowThreadProcessId(hwnd, ctypes.byref(p))
            if p.value != pid or not IsWindowVisible(hwnd):
                return True
            if win_class(hwnd) != "#32770":
                return True
            texts = [win_text(hwnd)]

            def child(ch, _x):
                t = win_text(ch)
                if t:
                    texts.append(t)
                return True

            EnumChildWindows(hwnd, EnumProc(child), 0)
            t = " || ".join(texts)
            if t not in dialogs:
                dialogs.append(t)
            return True

        before = len(dialogs)
        EnumWindows(EnumProc(cb), 0)
        if len(dialogs) > before:
            time.sleep(1.0)
            kill_excel()
            return
        time.sleep(0.25)


def test(idx, label, decl_code):
    dialogs.clear()
    kill_excel()
    time.sleep(1.5)
    work = os.path.join(temp, f"varblock_{idx}.xlsm")
    decl = os.path.join(temp, f"modDeclTest_{idx}.bas")
    probe = os.path.join(temp, f"modProbeTest_{idx}.bas")
    with open(decl, "w", encoding="utf-8", newline="\n") as f:
        f.write('Attribute VB_Name = "modDeclTest"\nOption Explicit\n\n' + decl_code + "\n")
    with open(probe, "w", encoding="utf-8", newline="\n") as f:
        f.write('Attribute VB_Name = "modProbe"\nOption Explicit\n\nPublic Function Probe1() As Boolean\n    Probe1 = mInReportRun\nEnd Function\n')

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    stop = threading.Event()
    out = ""
    try:
        wb = excel.Workbooks.Add()
        wb.SaveAs(work, FileFormat=52)
        vbp = wb.VBProject
        vbp.VBComponents.Import(decl)
        vbp.VBComponents.Import(probe)
        pid = wt.DWORD()
        GetWindowThreadProcessId(excel.Hwnd, ctypes.byref(pid))
        threading.Thread(target=watch, args=(pid.value, stop), daemon=True).start()
        try:
            v = excel.Run("Probe1")
            out = f"OK {v!r}"
        except Exception as e:
            out = "EXC " + str(e)[:70]
        time.sleep(1.2)
    finally:
        stop.set()
        pythoncom.CoUninitialize()
        kill_excel()
    bad = any("Compile error" in d for d in dialogs)
    print(f"{label:22s} {out}  compile_error={bad}", flush=True)
    return bad


lines = VARS.split("\n")
no_issues = [ln for ln in lines if "mAuditIssues" not in ln]
no_objects = [ln for ln in lines if "As Collection" not in ln and "As Object" not in ln]
only_objects = [ln for ln in lines if "As Collection" in ln or "As Object" in ln]
only_scalars = [ln for ln in lines if "As Collection" not in ln and "As Object" not in ln]

cases = [
    ("baseline", "Public mInReportRun As Boolean"),
    ("+issues", "Public mInReportRun As Boolean\nPublic mAuditIssues As Collection"),
    ("+repdict", "Public mInReportRun As Boolean\nPublic mRepSeenDict As Object"),
    ("+both", "Public mInReportRun As Boolean\nPublic mAuditIssues As Collection\nPublic mRepSeenDict As Object"),
    (
        "private+props",
        "Public mInReportRun As Boolean\n"
        "Private moIssues As Collection\n"
        "Private moRepSeen As Object\n\n"
        "Public Property Get AuditIssues() As Collection\n"
        "    Set AuditIssues = moIssues\n"
        "End Property\n\n"
        "Public Property Set AuditIssues(ByVal v As Collection)\n"
        "    Set moIssues = v\n"
        "End Property\n\n"
        "Public Property Get RepSeenDict() As Object\n"
        "    Set RepSeenDict = moRepSeen\n"
        "End Property\n\n"
        "Public Property Set RepSeenDict(ByVal v As Object)\n"
        "    Set moRepSeen = v\n"
        "End Property",
    ),
]
for i, (label, code) in enumerate(cases):
    test(i, label, code)
