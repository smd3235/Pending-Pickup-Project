Attribute VB_Name = "PendingPickupReport"
Option Explicit
Const ENABLE_ERL As Boolean = True

' ============================================================
' PENDING PICKUP REPORT BUILDER - v1.0
' Builds the Pending Pickup report from a current All Open Treatment Plans
' export and compares it against the prior month's Pending Pickup report.
' ============================================================

' ============================================================
' MAP OF THIS MODULE - snapshot of section starts and procedures.
'
'   #  SECTION                                      procedures
'
'   [  1] DECLARATIONS                              ENABLE_ERL, constants, mRunLogPath, mInReportRun
'   [  2] LOGGING                                   LogMsg, LogErr, BuildRunLogPath, WriteCompletionMarker
'   [  3] INTERACTIVE ENTRY                         SetupPendingPickupButton
'   [  4] CORE / MAIN BUILDER                       CreatePendingPickupReport
'   [  5] DASHBOARD                                 BuildDashboard, AddLogoFromTemplate, AddDashboardCard
'   [  6] MANAGER SHEET FORMATTING                  FormatManagerSheet
'   [  7] TABLE / ROW FORMATTING                    ApplyRepBanding, ApplyAgeWarningFont, HighlightMissingPickupDates, FormatYesNoColumns, SetManagerColumnWidths, FormatRawData, ApplyDataFormatting, StyleSectionHeader, StyleTableHeader, StyleDashboardTable
'   [  8] METRICS / AGING HELPERS                   CountUniqueManagerRepsOnSheet, CalculateAverageAge, CountNewManagerRows, CalculateAgingStats
'   [  9] DATE / TIME HELPERS                       NormalizeDateColumn, ApplySourceDateFormats, ApplyDateFormatsToOutput, ApplyDateFormatByHeader, TryGetExcelDate
'   [ 10] OUTPUT ARRAY / SORTING                    HighlightNewRecords, BuildOutputArray, SortManagerRowsBySalesRep, SortManagerColumns, ManagerColumnRank, BuildManagerOutputArray, ColumnLetter
'   [ 11] COLUMN / HEADER UTILITIES                 IsExcludedManagerHeader, FindHeaderColumnFromArray, CopyColumnFormats, MatchSourceColumnWidths, MatchSourceColumnWidthsForManager
'   [ 12] SHEET / WORKBOOK UTILITIES                FindHeaderRow, FindHeaderColumn, GetRecordKey, SafeValue, NormalizeText, LastUsedRow, LastUsedColumn, SheetExists, SafeSheetName, SortManagersByLastName, ManagerSortKey, SortTextArray
'   [ 13] WORKBOOK STATE                            EnsureMacroEnabledWorkbook, FindPriorSourceSheet
'   [ 14] DIALOG / FILE / FOLDER                    SelectExcelFile, SaveFolderSetting, GetSavedFolder, FindMostRecentPendingPickupReport, BuildOutputPath, GetSavedOutputFolder, SaveOutputFolder
' ============================================================


Private Const RAW_SHEET As String = "Raw Data"
Private Const MANAGER_HEADER As String = "Assigned Sales Manager"
Private Const STATUS_HEADER As String = "Treatment Status"
Private Const REP_HEADER As String = "Sales Representative"
Private Const NAME_HEADER As String = "Name"
Private Const KEY_HEADER As String = "EMPI"
Private Const PICKUP_DATE_HEADER As String = "AUDIT: TP End Date"
Private Const START_DATE_HEADER As String = "Start Date"
Private Const PENDING_TEXT As String = "Pending Pickup"
Private Const MACRO_VERSION As String = "1.0"
Private Const MANAGER_HEADER_ROW As Long = 16
Private Const LOG_PATH As String = "C:\Users\smd32\Downloads\Pending Pickup Project\log\macro-debug.log"
Private Const COMPLETION_MARKER_PATH As String = "C:\Users\smd32\Downloads\Pending Pickup Project\log\pending-pickup-complete.txt"
Private Const AUDIT_SHEET_NAME As String = "Audit Log"

' Audit tracking (populated during each report run)
Private mAuditStartTime As Date
Private mAuditSourcePath As String
Private mAuditOutputPath As String
Private mAuditTotalRecords As Long
Private mAuditPendingRecords As Long
Private mAuditNewRecords As Long
Private mAuditClearedRecords As Long
Private mAuditStillRecords As Long
Private mAuditMissingDates As Long
Private mAuditDuplicateKeys As Long
Private mAuditMissingManagers As Long
Private mAuditMissingStatus As Long
Private mAuditFutureDates As Long
Private mAuditVeryOldDates As Long
Private mAuditDuplicateNames As Long
Private mAuditErrorCount As Long
Private mAuditIssues As Collection
Private mRunLogPath As String
Private mInReportRun As Boolean
Private mRepSeenDict As Object

' ====================== LOGGING ======================
' Centralised debug / audit logging. Every run writes to LOG_PATH and, during a
' report run, also to the per-run log in mRunLogPath.
Private Sub LogMsg(ByVal msg As String)
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open LOG_PATH For Append As #f
    Print #f, Format(Now, "yyyy-mm-dd hh:n:ss") & " | " & msg
    Close #f

    If mInReportRun And Len(mRunLogPath) > 0 Then
        f = FreeFile
        Open mRunLogPath For Append As #f
        Print #f, Format(Now, "yyyy-mm-dd hh:n:ss") & " | " & msg
        Close #f
    End If
End Sub

Private Sub LogErr(ByVal procName As String)
    LogMsg "ERROR in " & procName & ": " & Err.Description & " (Err=" & Err.Number & ", Line=" & Erl & ")"
End Sub

Private Function BuildRunLogPath() As String
    Dim logDir As String
    On Error Resume Next
    logDir = Left$(LOG_PATH, InStrRev(LOG_PATH, "\"))
    BuildRunLogPath = logDir & "pending-pickup_" & Format(Now, "yyyy-mm-dd_HHMMSS") & ".log"
End Function

Private Sub WriteCompletionMarker(ByVal savePath As String)
    Dim f As Integer
    On Error Resume Next
    If Len(Dir$(Left$(COMPLETION_MARKER_PATH, InStrRev(COMPLETION_MARKER_PATH, "\")), vbDirectory)) = 0 Then MkDir Left$(COMPLETION_MARKER_PATH, InStrRev(COMPLETION_MARKER_PATH, "\"))
    f = FreeFile
    Open COMPLETION_MARKER_PATH For Output As #f
    Print #f, "completed=" & Format(Now, "yyyy-mm-dd hh:n:ss")
    Print #f, "saved=" & savePath
    Close #f
End Sub

' ====================== AUDIT ======================
' Centralised audit logging for each report run.
' Records metadata, record counts, data quality issues, and manager breakdowns
' both to the run log and to an Audit Log sheet in the output workbook.
Private Sub AuditInit(ByVal srcPath As String)
    mAuditStartTime = Now
    mAuditSourcePath = srcPath
    mAuditTotalRecords = 0
    mAuditPendingRecords = 0
    mAuditNewRecords = 0
    mAuditClearedRecords = 0
    mAuditStillRecords = 0
    mAuditMissingDates = 0
    mAuditDuplicateKeys = 0
    mAuditMissingManagers = 0
    mAuditMissingStatus = 0
    mAuditFutureDates = 0
    mAuditVeryOldDates = 0
    mAuditDuplicateNames = 0
    mAuditErrorCount = 0
    Set mAuditIssues = New Collection
End Sub

Private Sub AuditLogEvent(ByVal eventType As String, ByVal msg As String)
    On Error Resume Next
    Dim f As Integer
    f = FreeFile
    Open LOG_PATH For Append As #f
    Print #f, Format(Now, "yyyy-mm-dd hh:n:ss") & " | AUDIT | " & eventType & " | " & msg
    Close #f

    If mInReportRun And Len(mRunLogPath) > 0 Then
        f = FreeFile
        Open mRunLogPath For Append As #f
        Print #f, Format(Now, "yyyy-mm-dd hh:n:ss") & " | AUDIT | " & eventType & " | " & msg
        Close #f
    End If

    If Len(msg) > 0 Then
        mAuditIssues.Add eventType & ": " & msg
    End If
End Sub

Private Sub AuditIncrementCounter(ByVal counterName As String)
    Select Case counterName
        Case "TotalRecords": mAuditTotalRecords = mAuditTotalRecords + 1
        Case "PendingRecords": mAuditPendingRecords = mAuditPendingRecords + 1
        Case "NewRecords": mAuditNewRecords = mAuditNewRecords + 1
        Case "ClearedRecords": mAuditClearedRecords = mAuditClearedRecords + 1
        Case "StillRecords": mAuditStillRecords = mAuditStillRecords + 1
        Case "MissingDates": mAuditMissingDates = mAuditMissingDates + 1
        Case "DuplicateKeys": mAuditDuplicateKeys = mAuditDuplicateKeys + 1
        Case "MissingManagers": mAuditMissingManagers = mAuditMissingManagers + 1
        Case "Errors": mAuditErrorCount = mAuditErrorCount + 1
    End Select
End Sub

Private Sub AuditWriteSheet(ByVal wb As Workbook)
    Dim ws As Worksheet, r As Long
    On Error Resume Next

    ' Remove existing audit sheet if present
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets(AUDIT_SHEET_NAME).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = AUDIT_SHEET_NAME

    ' Title
    ws.Cells(1, 1).Value = "PENDING PICKUP REPORT - AUDIT LOG"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    ws.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    ws.Cells(1, 1).Interior.Color = RGB(31, 78, 121)
    ws.Cells(1, 1).HorizontalAlignment = xlCenter

    ' Timestamp row
    ws.Cells(3, 1).Value = "Report Generated:"
    ws.Cells(3, 2).Value = Format(mAuditStartTime, "yyyy-mm-dd hh:n:ss")
    ws.Cells(4, 1).Value = "Source File:"
    ws.Cells(4, 2).Value = mAuditSourcePath
    ws.Cells(5, 1).Value = "Output File:"
    ws.Cells(5, 2).Value = mAuditOutputPath
    ws.Cells(6, 1).Value = "Macro Version:"
    ws.Cells(6, 2).Value = MACRO_VERSION

    ' Record counts section
    ws.Cells(8, 1).Value = "RECORD COUNTS"
    ws.Cells(8, 1).Font.Bold = True
    ws.Cells(8, 1).Interior.Color = RGB(31, 78, 121)
    ws.Cells(8, 1).Font.Color = RGB(255, 255, 255)
    ws.Cells(8, 1).Resize(1, 2).Merge

    ws.Cells(9, 1).Value = "Total Records in Source:"
    ws.Cells(9, 2).Value = mAuditTotalRecords
    ws.Cells(10, 1).Value = "Pending Pickup Records:"
    ws.Cells(10, 2).Value = mAuditPendingRecords
    ws.Cells(11, 1).Value = "New Since Prior Report:"
    ws.Cells(11, 2).Value = mAuditNewRecords
    ws.Cells(12, 1).Value = "Cleared Since Prior Report:"
    ws.Cells(12, 2).Value = mAuditClearedRecords
    ws.Cells(13, 1).Value = "Still Pending:"
    ws.Cells(13, 2).Value = mAuditStillRecords

    ' Data quality section
    ws.Cells(15, 1).Value = "DATA QUALITY ISSUES"
    ws.Cells(15, 1).Font.Bold = True
    ws.Cells(15, 1).Interior.Color = RGB(31, 78, 121)
    ws.Cells(15, 1).Font.Color = RGB(255, 255, 255)
    ws.Cells(15, 1).Resize(1, 2).Merge

    ws.Cells(16, 1).Value = "Missing Pickup Dates:"
    ws.Cells(16, 2).Value = mAuditMissingDates
    ws.Cells(17, 1).Value = "Duplicate Keys (EMPI):"
    ws.Cells(17, 2).Value = mAuditDuplicateKeys
    ws.Cells(18, 1).Value = "Missing Manager Names:"
    ws.Cells(18, 2).Value = mAuditMissingManagers
    ws.Cells(19, 1).Value = "Errors Encountered:"
    ws.Cells(19, 2).Value = mAuditErrorCount
    ws.Cells(20, 1).Value = "Missing Status Values:"
    ws.Cells(20, 2).Value = mAuditMissingStatus
    ws.Cells(21, 1).Value = "Future Pickup Dates:"
    ws.Cells(21, 2).Value = mAuditFutureDates
    ws.Cells(22, 1).Value = "Very Old Dates (>5 years):"
    ws.Cells(22, 2).Value = mAuditVeryOldDates
    ws.Cells(23, 1).Value = "Duplicate Rep Names:"
    ws.Cells(23, 2).Value = mAuditDuplicateNames

    ' Audit details section
    r = 21
    ws.Cells(r, 1).Value = "AUDIT DETAILS"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Interior.Color = RGB(31, 78, 121)
    ws.Cells(r, 1).Font.Color = RGB(255, 255, 255)
    ws.Cells(r, 1).Resize(1, 2).Merge

    r = r + 1
    ws.Cells(r, 1).Value = "Issue Type"
    ws.Cells(r, 2).Value = "Description"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 1).Interior.Color = RGB(217, 234, 247)
    ws.Cells(r, 2).Interior.Color = RGB(217, 234, 247)
    r = r + 1

    Dim item As Variant
    For Each item In mAuditIssues
        ws.Cells(r, 1).Value = Split(item, ":")(0)
        ws.Cells(r, 2).Value = Mid(item, InStr(item, ":") + 2)
        r = r + 1
    Next item

    ' Auto-fit columns
    ws.Columns(1).AutoFit
    ws.Columns(2).AutoFit

    LogMsg "Audit sheet written"
End Sub

' ====================== INTERACTIVE ENTRY ======================
' Owns the START HERE canvas button and its wiring.
Private Sub SetupPendingPickupButton()
    Dim ws As Worksheet, shp As Shape, btn As Shape
    On Error GoTo ButtonFail

    If Not EnsureMacroEnabledWorkbook() Then Exit Sub

    Set ws = ThisWorkbook.Worksheets("START HERE")

    For Each shp In ws.Shapes
        If shp.Name = "btnCreatePendingPickupReport" Then
            shp.OnAction = "CreatePendingPickupReport"
            MsgBox "The CREATE PENDING PICKUP REPORT button is already installed.", vbInformation
            Exit Sub
        End If
    Next shp

    Set btn = ws.Shapes.AddShape(5, ws.Range("B14").Left + 20, ws.Range("B14").Top + 8, 360, 48)   ' 5 = rounded rectangle
    btn.Name = "btnCreatePendingPickupReport"
    btn.Fill.ForeColor.RGB = RGB(31, 78, 121)
    btn.Line.ForeColor.RGB = RGB(246, 206, 62)
    btn.Line.Weight = 1.5
    With btn.TextFrame2.TextRange
        .Text = "CREATE PENDING PICKUP REPORT"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Name = "Segoe UI"
        .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
    btn.TextFrame2.VerticalAnchor = 3
    btn.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    btn.OnAction = "CreatePendingPickupReport"
    MsgBox "Button installed. Click it whenever you need the report. The macro will ask you to select the current source file and prior report, and will remember your report folders.", vbInformation
    Exit Sub

ButtonFail:
    MsgBox "I could not install the button." & vbCrLf & Err.Description, vbCritical
End Sub

' ====================== CORE / MAIN BUILDER ======================
' Owns the full report build: source selection, prior comparison, manager tabs,
' dashboard, Raw Data copy, and save. Public entry point wired to the button.
Public Function CreatePendingPickupReport() As Boolean
    CreatePendingPickupReport = False
    #If ENABLE_ERL Then
    On Error GoTo CleanFail
    #End If
    Dim srcWb As Workbook, sourceWb As Workbook, raw As Worksheet, prior As Worksheet, priorWb As Workbook, outWb As Workbook
    Dim currentData As Variant, priorData As Variant
    Dim currentRows As Collection, priorRows As Collection
    Dim currentByKey As Object, priorByKey As Object
    Dim managerRows As Object, managerNames As Object
    Dim headerRow As Long, priorHeaderRow As Long
    Dim lastRow As Long, lastCol As Long, priorLastRow As Long, priorLastCol As Long
    Dim managerCol As Long, statusCol As Long, repCol As Long, nameCol As Long, pickupDateCol As Long, startDateCol As Long
    Dim priorStatusCol As Long, priorNameCol As Long, priorManagerCol As Long
    Dim r As Long, i As Long
    Dim manager As String, status As String
    Dim ws As Worksheet, dashboard As Worksheet, rawOut As Worksheet
    Dim arrOut As Variant, arrRaw As Variant
    Dim managers As Variant
    Dim savePath As String
    Dim priorPath As Variant
    Dim oldCalc As XlCalculation, oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean
    Dim hasPrior As Boolean
    Dim currentCount As Long, priorCount As Long, newCount As Long, clearedCount As Long, stillCount As Long
    Dim age30Count As Long, age60Count As Long, age90Count As Long, missingDateCount As Long
    Dim errText As String
    Dim managerLastCol As Long, managerPickupDateCol As Long, managerStartDateCol As Long, managerAgeCol As Long, managerRepCol As Long
    Dim sortedManagerRows As Collection
    Dim reportStamp As String, sourceName As String
    #If Not ENABLE_ERL Then
    On Error GoTo CleanFail
    #End If
    mInReportRun = True
    mRunLogPath = BuildRunLogPath()
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    LogMsg "=== CreatePendingPickupReport started ==="
    LogMsg "Run log: " & mRunLogPath
    LogMsg "Completion marker: " & COMPLETION_MARKER_PATH
    LogMsg "Excel version: " & Application.Version
    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts

    ' Initialize audit tracking
    AuditInit ThisWorkbook.FullName
    AuditLogEvent "START", "Report run started"

    Set mRepSeenDict = CreateObject("Scripting.Dictionary")
    mRepSeenDict.CompareMode = vbTextCompare

    Set srcWb = ThisWorkbook

    ' Always perform the format check before doing any report work.
    ' The master must be .xlsm so the imported VBA module is retained.
    If Not EnsureMacroEnabledWorkbook() Then GoTo CleanExit
    Set srcWb = ThisWorkbook
    LogMsg "EnsureMacroEnabledWorkbook OK"

    ' Select the CURRENT source report. Nothing needs to be pasted into the master.
    Application.StatusBar = "Select the current All Open Treatment Plans report..."
    Dim currentPath As Variant
    currentPath = SelectExcelFile( _
        "Select the CURRENT All Open Treatment Plans report", _
        GetSavedFolder("CurrentFolder"), _
        "", _
        False)
    If currentPath = False Then
        LogMsg "No current file selected"
        MsgBox "No current source file was selected. The report was not run.", vbExclamation
        GoTo CleanExit
    End If
    If StrComp(CStr(currentPath), srcWb.FullName, vbTextCompare) = 0 Then
        MsgBox "Please select the current source report, not the Pending Pickup master workbook.", vbExclamation
        GoTo CleanExit
    End If
    SaveFolderSetting "CurrentFolder", CStr(currentPath)
    LogMsg "CURRENT selected: " & currentPath

    Application.StatusBar = "Opening current source report..."
    Set sourceWb = Workbooks.Open(Filename:=CStr(currentPath), ReadOnly:=True, UpdateLinks:=False)
    reportStamp = Format(Now, "mm/dd/yyyy h:mm AM/PM")
    sourceName = sourceWb.Name
    Set raw = FindPriorSourceSheet(sourceWb)
    If raw Is Nothing Then
        LogMsg "FindPriorSourceSheet returned Nothing"
        MsgBox "I could not find the required headers in the selected current source file." & vbCrLf & vbCrLf & _
               "Required: " & MANAGER_HEADER & ", " & STATUS_HEADER & ", " & REP_HEADER & ", " & NAME_HEADER, vbCritical
        GoTo CleanExit
    End If
    LogMsg "Source sheet: " & raw.Name

    headerRow = FindHeaderRow(raw)
    If headerRow = 0 Then
        LogMsg "FindHeaderRow returned 0"
        MsgBox "I could not find the required headers on Raw Data:" & vbCrLf & _
               "• " & MANAGER_HEADER & vbCrLf & "• " & STATUS_HEADER & vbCrLf & _
               "• " & REP_HEADER & vbCrLf & "• " & NAME_HEADER, vbCritical
        GoTo CleanExit
    End If
    LogMsg "Header row: " & headerRow

    managerCol = FindHeaderColumn(raw, headerRow, MANAGER_HEADER)
    statusCol = FindHeaderColumn(raw, headerRow, STATUS_HEADER)
    repCol = FindHeaderColumn(raw, headerRow, REP_HEADER)
    nameCol = FindHeaderColumn(raw, headerRow, NAME_HEADER)

    pickupDateCol = FindHeaderColumn(raw, headerRow, PICKUP_DATE_HEADER)
    startDateCol = FindHeaderColumn(raw, headerRow, START_DATE_HEADER)

    lastRow = LastUsedRow(raw)
    lastCol = LastUsedColumn(raw)
    LogMsg "lastRow=" & lastRow & " lastCol=" & lastCol
    If lastRow <= headerRow Then
        MsgBox "There are no current data rows to process.", vbInformation
        GoTo CleanExit
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Preparing current Pending Pickup data..."

    ' Convert recognizable text dates to real Excel dates and apply date formats.
    LogMsg "NormalizeDateColumn start"
    NormalizeDateColumn raw, headerRow, lastRow, startDateCol
    NormalizeDateColumn raw, headerRow, lastRow, pickupDateCol
    LogMsg "NormalizeDateColumn done"

    LogMsg "Value2 start"
    currentData = raw.Range(raw.Cells(headerRow, 1), raw.Cells(lastRow, lastCol)).Value2
    LogMsg "Value2 done rows=" & UBound(currentData, 1)
    LogMsg "Source dimensions: rows=" & lastRow & " cols=" & lastCol & " statusCol=" & statusCol & " managerCol=" & managerCol & " pickupDateCol=" & pickupDateCol
    Set managerRows = CreateObject("Scripting.Dictionary")
    managerRows.CompareMode = vbTextCompare
    Set managerNames = CreateObject("Scripting.Dictionary")
    managerNames.CompareMode = vbTextCompare
    Set currentRows = New Collection

    For r = 2 To UBound(currentData, 1)
        On Error Resume Next
        status = Trim$(CStr(currentData(r, statusCol)))
        If Err.Number <> 0 Then
            AuditLogEvent "ERROR", "Row " & r & ": Cannot read status cell - " & Err.Number & " " & Err.Description
            Err.Clear
        ElseIf Len(Trim$(status)) = 0 Then
            mAuditMissingStatus = mAuditMissingStatus + 1
            AuditLogEvent "MISSING_STATUS", "Row " & r & ": Missing status value"
        ElseIf InStr(1, status, PENDING_TEXT, vbTextCompare) > 0 Then
            currentRows.Add r

            manager = Trim$(CStr(currentData(r, managerCol)))
            If Len(manager) = 0 Then
                manager = "Unassigned"
                mAuditMissingManagers = mAuditMissingManagers + 1
                AuditLogEvent "MISSING_MANAGER", "Row " & r & ": Missing manager name"
            End If
            If Not managerRows.Exists(manager) Then
                Set managerRows(manager) = New Collection
                managerNames.Add manager, manager
            End If
            managerRows(manager).Add r

            ' Audit: check for future dates
            If Not IsEmpty(currentData(r, pickupDateCol)) Then
                Dim dtCheck As Date
                If TryGetExcelDate(currentData(r, pickupDateCol), dtCheck) Then
                    If dtCheck > Date Then
                        mAuditFutureDates = mAuditFutureDates + 1
                        AuditLogEvent "FUTURE_DATE", "Row " & r & ": Future pickup date " & Format(dtCheck, "yyyy-mm-dd")
                    ElseIf DateDiff("yyyy", dtCheck, Date) > 5 Then
                        mAuditVeryOldDates = mAuditVeryOldDates + 1
                        AuditLogEvent "VERY_OLD_DATE", "Row " & r & ": Very old pickup date " & Format(dtCheck, "yyyy-mm-dd") & " (" & DateDiff("yyyy", dtCheck, Date) & " years ago)"
                    End If
                End If
            End If

            ' Audit: check for duplicate names across managers
            Dim repName As String
            repName = Trim$(CStr(currentData(r, repCol)))
            If Len(repName) > 0 Then
                If repSeenDict.Exists(repName) Then
                    mAuditDuplicateNames = mAuditDuplicateNames + 1
                    AuditLogEvent "DUPLICATE_NAME", "Row " & r & ": Duplicate rep name '" & repName & "' (first seen at row " & repSeenDict(repName) & ")"
                Else
                    repSeenDict.Add repName, r
                End If
            End If
        End If
    Next r
    On Error GoTo CleanFail

    currentCount = currentRows.Count
    LogMsg "currentCount=" & currentCount & " managers=" & managerNames.Count
    AuditLogEvent "COUNTS", "currentCount=" & currentCount & " managers=" & managerNames.Count & " totalRows=" & UBound(currentData, 1)
    mAuditTotalRecords = UBound(currentData, 1) - 1 ' exclude header
    mAuditPendingRecords = currentCount
    If currentCount = 0 Then
        MsgBox "No rows were found where Treatment Status contains '" & PENDING_TEXT & "'.", vbInformation
        GoTo CleanExit
    End If

    Set priorRows = New Collection
    Set priorByKey = CreateObject("Scripting.Dictionary")
    priorByKey.CompareMode = vbTextCompare

    ' Select the prior report file. Cancel means current-only reporting.
    Application.StatusBar = "Select the prior Pending Pickup report..."
    Dim priorSuggested As String
    priorSuggested = FindMostRecentPendingPickupReport(GetSavedFolder("PriorFolder"), srcWb.FullName)
    LogMsg "PRIOR picker (suggested: " & priorSuggested & ")"
    priorPath = SelectExcelFile( _
        "Select the PRIOR Pending Pickup report (Cancel for no comparison)", _
        GetSavedFolder("PriorFolder"), _
        priorSuggested, _
        True)

    hasPrior = False
    LogMsg "priorPath=False? " & (priorPath = False)
    If priorPath <> False Then
        LogMsg "priorPath selected: " & priorPath
        SaveFolderSetting "PriorFolder", CStr(priorPath)
        If StrComp(CStr(priorPath), srcWb.FullName, vbTextCompare) = 0 Then
            MsgBox "The selected prior report is the current master workbook. Please select the older report file instead.", vbExclamation
        Else
            Application.StatusBar = "Opening prior report..."
            Set priorWb = Workbooks.Open(Filename:=CStr(priorPath), ReadOnly:=True, UpdateLinks:=False)
            Set prior = FindPriorSourceSheet(priorWb)
            If prior Is Nothing Then
                MsgBox "I could not find a worksheet in the selected prior file containing the required headers." & vbCrLf & vbCrLf & _
                       "The prior report was not used for comparison.", vbExclamation
            Else
                priorHeaderRow = FindHeaderRow(prior)
                priorStatusCol = FindHeaderColumn(prior, priorHeaderRow, STATUS_HEADER)
                priorNameCol = FindHeaderColumn(prior, priorHeaderRow, NAME_HEADER)

                priorManagerCol = FindHeaderColumn(prior, priorHeaderRow, MANAGER_HEADER)
                priorLastRow = LastUsedRow(prior)
                priorLastCol = LastUsedColumn(prior)
                If priorLastRow > priorHeaderRow Then hasPrior = True
            End If
        End If
    End If

    If hasPrior Then
        Application.StatusBar = "Reading prior Pending Pickup data..."
        priorData = prior.Range(prior.Cells(priorHeaderRow, 1), prior.Cells(priorLastRow, priorLastCol)).Value2
        For r = 2 To UBound(priorData, 1)
            status = Trim$(CStr(priorData(r, priorStatusCol)))
            If InStr(1, status, PENDING_TEXT, vbTextCompare) > 0 Then
                priorRows.Add r

                If Not priorByKey.Exists(key) Then priorByKey.Add key, r
            End If
        Next r
    End If

    priorCount = priorRows.Count
    stillCount = 0
    newCount = 0
    clearedCount = 0
    LogMsg "priorCount=" & priorCount & " hasPrior=" & hasPrior

    If hasPrior Then


                stillCount = stillCount + 1
            Else
                newCount = newCount + 1

            End If
        Next key

            If Not currentByKey.Exists(key) Then clearedCount = clearedCount + 1
        Next key
    End If

    overallAvgAge = CalculateAverageAge(currentData, currentRows, pickupDateCol)

    If hasPrior Then
        For i = 1 To currentRows.Count
            r = CLng(currentRows(i))


                manager = Trim$(CStr(currentData(r, managerCol)))
                If Len(manager) = 0 Then manager = "Unassigned"


                Else

                End If
            End If
        Next i
        If priorManagerCol > 0 Then
            For i = 1 To priorRows.Count
                r = CLng(priorRows(i))


                    manager = Trim$(CStr(priorData(r, priorManagerCol)))
                    If Len(manager) = 0 Then manager = "Unassigned"


                    Else

                    End If
                End If
            Next i
        End If
    End If

    CalculateAgingStats currentData, currentRows, pickupDateCol, age30Count, age60Count, age90Count, missingDateCount
    LogMsg "Aging: 30=" & age30Count & " 60=" & age60Count & " 90=" & age90Count & " missing=" & missingDateCount
    AuditLogEvent "AGING", "30=" & age30Count & " 60=" & age60Count & " 90=" & age90Count & " missing=" & missingDateCount
    mAuditMissingDates = missingDateCount

    raw.Range(raw.Cells(headerRow, 1), raw.Cells(lastRow, lastCol)).Columns.AutoFit
    ApplySourceDateFormats raw, headerRow, lastRow, startDateCol, pickupDateCol

    Set outWb = Workbooks.Add(xlWBATWorksheet)
    Set dashboard = outWb.Worksheets(1)
    dashboard.Name = "Dashboard"
    LogMsg "Output workbook created"

    managers = managerNames.Keys
    If managerNames.Count > 1 Then SortManagersByLastName managers

    For i = LBound(managers) To UBound(managers)
        manager = CStr(managers(i))
        LogMsg "Creating manager tab: " & manager
        Application.StatusBar = "Creating manager tab: " & manager
        Set ws = outWb.Worksheets.Add(After:=outWb.Worksheets(outWb.Worksheets.Count))
        ws.Name = SafeSheetName(manager, outWb)

        Set sortedManagerRows = SortManagerRowsBySalesRep(currentData, managerRows(manager), repCol, nameCol)
        arrOut = BuildManagerOutputArray(currentData, sortedManagerRows, lastCol, pickupDateCol)
        managerLastCol = UBound(arrOut, 2)
        ws.Range(ws.Cells(MANAGER_HEADER_ROW, 1), ws.Cells(MANAGER_HEADER_ROW + UBound(arrOut, 1) - 1, managerLastCol)).Value = arrOut
        managerPickupDateCol = FindHeaderColumn(ws, MANAGER_HEADER_ROW, PICKUP_DATE_HEADER)
        managerStartDateCol = FindHeaderColumn(ws, MANAGER_HEADER_ROW, START_DATE_HEADER)
        managerAgeCol = FindHeaderColumn(ws, MANAGER_HEADER_ROW, "Pickup Age (Days)")
        managerRepCol = FindHeaderColumn(ws, MANAGER_HEADER_ROW, REP_HEADER)
        FormatManagerSheet ws, manager, managerRows\(manager\).Count, managerLastCol, managerPickupDateCol, managerStartDateCol, managerAgeCol, managerRepCol, currentData, sortedManagerRows, hasPrior
        LogMsg "  ApplyDateFormatsToOutput: " & manager
        ApplyDateFormatsToOutput ws, MANAGER_HEADER_ROW, LastUsedRow(ws), managerStartDateCol, managerPickupDateCol
        LogMsg "  MatchSourceColumnWidthsForManager: " & manager
        MatchSourceColumnWidthsForManager raw, ws, lastCol
        LogMsg "  manager tab done: " & manager
    Next i
    LogMsg "All manager tabs done"

    ' Build the dashboard after manager sheets exist so manager names can link directly to their tabs.
    LogMsg "BuildDashboard start"
    BuildDashboard dashboard, currentCount, priorCount, newCount, clearedCount, stillCount, _
                   age30Count, age60Count, age90Count, missingDateCount, hasPrior, managers, managerRows, managerSheetNames, currentData, keyCol, nameCol, pickupDateCol, _
                   reportStamp, sourceName, currentRows
    LogMsg "BuildDashboard done"

    Set rawOut = outWb.Worksheets.Add(After:=outWb.Worksheets(outWb.Worksheets.Count))
    rawOut.Name = "Raw Data"
    LogMsg "Raw Data sheet created"
    arrRaw = BuildOutputArray(currentData, currentRows, lastCol)
    rawOut.Range(rawOut.Cells(1, 1), rawOut.Cells(UBound(arrRaw, 1), lastCol)).Value = arrRaw
    FormatRawData rawOut, lastCol, UBound(arrRaw, 1)
    CopyColumnFormats raw, rawOut, lastCol
    MatchSourceColumnWidths raw, rawOut, lastCol
    ApplyDateFormatsToOutput rawOut, 1, UBound(arrRaw, 1), startDateCol, pickupDateCol
    ApplyDateFormatByHeader rawOut, 1, UBound(arrRaw, 1), "Patient Date of Birth"
    ApplyDateFormatByHeader rawOut, 1, UBound(arrRaw, 1), "Last Modified"

    Application.StatusBar = "Saving finished Pending Pickup report..."
    savePath = BuildOutputPath(srcWb)
    LogMsg "Saving to: " & savePath

    ' Hide unused rows/columns on every sheet so the report looks clean
    Dim hideWs As Worksheet
    For Each hideWs In outWb.Worksheets
        HideUnusedCells hideWs
    Next hideWs

    ' Write audit data to the output workbook
    mAuditOutputPath = savePath
    AuditWriteSheet outWb
    AuditLogEvent "SAVE", "Report saved to: " & savePath

    outWb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    LogMsg "Save done"
    WriteCompletionMarker savePath
    LogMsg "Completion marker written: " & COMPLETION_MARKER_PATH
    LogMsg "=== CreatePendingPickupReport completed ==="
    CreatePendingPickupReport = True

    Application.StatusBar = False
    Exit Function

CleanExit:
    On Error Resume Next
    mInReportRun = False
    mRunLogPath = ""
    If Not priorWb Is Nothing Then priorWb.Close SaveChanges:=False
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
    On Error GoTo 0
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc
    CreatePendingPickupReport = False
    Exit Function

CleanFail:
    errText = Err.Description & " (at line " & Erl & ")"
    LogMsg "FAILED: " & errText & " | Err=" & Err.Number
    On Error Resume Next
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.Calculation = oldCalc
    mInReportRun = False
    mRunLogPath = ""
    If Not outWb Is Nothing Then outWb.Close SaveChanges:=False
    If Not priorWb Is Nothing Then priorWb.Close SaveChanges:=False
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
    On Error GoTo 0
    LogMsg "=== CleanFail exit ==="
    CreatePendingPickupReport = False
    MsgBox "The Pending Pickup report could not be created." & vbCrLf & errText, vbCritical
End Function


' ====================== DASHBOARD ======================
' Builds the Dashboard tab: KPI cards, manager table, charts, and logo.
Private Sub BuildDashboard(ws As Worksheet, currentCount As Long, priorCount As Long, newCount As Long, clearedCount As Long, stillCount As Long, _
    On Error GoTo LocalFail
    LocalFail:
        LogErr "BuildDashboard"
        Err.Clear
        Resume Next

    rg.Interior.Color = fillColor
    rg.Borders.LineStyle = xlContinuous
    rg.Borders.Weight = xlMedium
    With rg.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(246, 206, 62)
    End With

    ' Label occupies the top rows; the value gets its own merged bottom row so
    ' the value never overwrites the label.
    Set labelRg = rg.Resize(rg.Rows.Count - 1, rg.Columns.Count)
    Set valueRow = rg.Offset(rg.Rows.Count - 1, 0).Resize(1, rg.Columns.Count)
    labelRg.Merge
    labelRg.Value = labelText & vbCrLf & subText
    With labelRg
        .Font.Bold = True
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    valueRow.Merge
    valueRow.Value = CStr(valueText)
    With valueRow
        .Font.Bold = True
        .Font.Name = "Segoe UI"
        .Font.Size = 22
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
End Sub

' ====================== MANAGER SHEET FORMATTING ======================
' Formats one manager tab: KPI cards, data table, banding, highlighting, age warnings.
Private Sub FormatManagerSheet(ws As Worksheet, managerName As String, pendingCount As Long, managerLastCol As Long, managerPickupDateCol As Long, managerStartDateCol As Long, managerAgeCol As Long, managerRepCol As Long, _
    On Error GoTo LocalFail
    LocalFail:
        LogErr "FormatManagerSheet"
        Err.Clear
        Resume Next

    For r = firstRow To lastRow
        v = ws.Cells(r, ageCol).Value
        On Error Resume Next
        If IsNumeric(v) And Len(Trim$(CStr(v))) > 0 Then
            If CDbl(v) > 30 Then ws.Cells(r, ageCol).Font.Color = RGB(192, 0, 0) Else ws.Cells(r, ageCol).Font.Color = RGB(0, 0, 0)
        Else
            ws.Cells(r, ageCol).Font.Color = RGB(0, 0, 0)
        End If
        On Error GoTo 0
    Next r
End Sub

Private Sub HighlightMissingPickupDates(ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal pickupDateCol As Long)
    Dim r As Long, v As Variant, dt As Date
    If pickupDateCol <= 0 Or firstRow > lastRow Then Exit Sub
    On Error Resume Next
    For r = firstRow To lastRow
        v = ws.Cells(r, pickupDateCol).Value
        If Len(Trim$(CStr(v))) = 0 Or Not TryGetExcelDate(v, dt) Then
            With ws.Cells(r, pickupDateCol)
                .Interior.Color = RGB(255, 199, 206)
                .Font.Bold = True
                .Font.Color = RGB(156, 0, 6)
            End With
        End If
    Next r
    On Error GoTo 0
End Sub

Private Sub FormatYesNoColumns(ws As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long, ByVal lastCol As Long)
    Dim c As Long, r As Long, h As String, v As String
    If lastRow <= headerRow Then Exit Sub
    On Error Resume Next
    For c = 1 To lastCol
        h = LCase$(Trim$(CStr(ws.Cells(headerRow, c).Value)))
        If InStr(1, h, "no rental equipment provided", vbTextCompare) > 0 Then
            For r = headerRow + 1 To lastRow
                v = LCase$(Trim$(CStr(ws.Cells(r, c).Value)))
                If Len(v) > 0 Then
                    If v = "yes" Or v = "y" Or v = "true" Then
                        ws.Cells(r, c).Value = "Yes"
                        ws.Cells(r, c).Font.Color = RGB(0, 97, 0)
                    ElseIf v = "no" Or v = "n" Or v = "false" Then
                        ws.Cells(r, c).Value = "No"
                        ws.Cells(r, c).Font.Color = RGB(156, 0, 6)
                    Else
                        ws.Cells(r, c).Value = "Yes"
                        ws.Cells(r, c).Font.Color = RGB(0, 97, 0)
                    End If
                Else
                    ws.Cells(r, c).Value = "No"
                    ws.Cells(r, c).Font.Color = RGB(156, 0, 6)
                End If
            Next r
        End If
    Next c
    On Error GoTo 0
End Sub

Private Sub SetManagerColumnWidths(ws As Worksheet, lastCol As Long)
    Dim c As Long, headerText As String, minWidth As Double
    If lastCol <= 0 Then Exit Sub
    ws.Range(ws.Cells(MANAGER_HEADER_ROW, 1), ws.Cells(LastUsedRow(ws), lastCol)).Columns.AutoFit
    For c = 1 To lastCol
        headerText = LCase$(Trim$(CStr(ws.Cells(MANAGER_HEADER_ROW, c).Value)))
        minWidth = 16
        Select Case headerText
            Case LCase$(REP_HEADER): minWidth = 24
            Case LCase$(START_DATE_HEADER): minWidth = 14
            Case LCase$(NAME_HEADER): minWidth = 26
            Case LCase$("Billing Customer"): minWidth = 26
            Case LCase$("Prescribing Physician"): minWidth = 28
            Case LCase$("No Rental Equipment Provided"): minWidth = 26
            Case LCase$(PICKUP_DATE_HEADER): minWidth = 19
            Case LCase$("Pickup Age (Days)"): minWidth = 16
        End Select
        If ws.Columns(c).ColumnWidth < minWidth Then ws.Columns(c).ColumnWidth = minWidth
    Next c
End Sub

Private Sub FormatRawData(ws As Worksheet, lastCol As Long, lastRow As Long)
    ApplyDataFormatting ws, 1, lastRow, lastCol
    ws.Tab.Color = RGB(128, 128, 128)
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "1:1"
    End With
    ws.Activate
    On Error Resume Next
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 85
    On Error GoTo 0
End Sub

Private Sub ApplyDataFormatting(ws As Worksheet, headerRow As Long, lastRow As Long, lastCol As Long)
    On Error GoTo LocalFail
    LocalFail:
        LogErr "ApplyDataFormatting"
        Err.Clear
        Resume Next

    rg.Font.Name = "Segoe UI"
    rg.Borders.LineStyle = xlContinuous
    rg.HorizontalAlignment = xlCenter
    With rg.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(246, 206, 62)
    End With
End Sub

Private Sub StyleDashboardTable(rg As Range)
    Dim r As Long
    rg.Borders.LineStyle = xlContinuous
    rg.Borders.Weight = xlHairline
    For r = 1 To rg.Rows.Count
        If r Mod 2 = 0 Then rg.Rows(r).Interior.Color = RGB(242, 242, 242)
    Next r
End Sub

' ====================== METRICS / AGING HELPERS ======================
' Manager stats: unique rep count, average age, new-row count, aging bucket stats.
Private Function CountUniqueManagerRepsOnSheet(ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal repCol As Long) As Long
    Dim d As Object, r As Long, repName As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    If repCol <= 0 Or firstRow > lastRow Then Exit Function
    For r = firstRow To lastRow
        repName = NormalizeText(CStr(ws.Cells(r, repCol).Value))
        If Len(repName) > 0 Then
            If Not d.Exists(repName) Then d.Add repName, True
        End If
    Next r
    CountUniqueManagerRepsOnSheet = d.Count
End Function

Private Function CountRepsInSheet(ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal repCol As Long) As Object
    Dim d As Object, r As Long, repName As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    If repCol <= 0 Or firstRow > lastRow Then Exit Function
    For r = firstRow To lastRow
        repName = NormalizeText(CStr(ws.Cells(r, repCol).Value))
        If Len(repName) > 0 Then
            If Not d.Exists(repName) Then d.Add repName, 0
            d(repName) = d(repName) + 1
        End If
    Next r
    Set CountRepsInSheet = d
End Function

Private Function CalculateAverageAge(ByRef data As Variant, ByVal rowsToUse As Collection, ByVal pickupDateCol As Long) As Variant
    On Error GoTo LocalFail
    LocalFail:
        LogErr "CalculateAverageAge"
        Err.Clear
        Resume Next

        LogErr "CalculateAgingStats"
        Err.Clear
        Resume Next

    Dim r As Long, dt As Date
    If dateCol <= 0 Or lastRow <= headerRow Then Exit Sub
    On Error Resume Next
    For r = headerRow + 1 To lastRow
        If TryGetExcelDate(ws.Cells(r, dateCol).Value, dt) Then
            ws.Cells(r, dateCol).Value = dt
        End If
    Next r
    ws.Range(ws.Cells(headerRow + 1, dateCol), ws.Cells(lastRow, dateCol)).NumberFormat = "mm/dd/yyyy"
    On Error GoTo 0
End Sub

Private Sub ApplySourceDateFormats(ws As Worksheet, headerRow As Long, lastRow As Long, startDateCol As Long, pickupDateCol As Long)
    If startDateCol > 0 And lastRow > headerRow Then ws.Range(ws.Cells(headerRow + 1, startDateCol), ws.Cells(lastRow, startDateCol)).NumberFormat = "mm/dd/yyyy"
    If pickupDateCol > 0 And lastRow > headerRow Then ws.Range(ws.Cells(headerRow + 1, pickupDateCol), ws.Cells(lastRow, pickupDateCol)).NumberFormat = "mm/dd/yyyy"
End Sub

Private Sub ApplyDateFormatsToOutput(ws As Worksheet, headerRow As Long, lastRow As Long, startDateCol As Long, pickupDateCol As Long)
    On Error GoTo LocalFail
    LocalFail:
        LogErr "ApplyDateFormatsToOutput"
        Err.Clear
        Resume Next

        dt = CDate(v)
    ElseIf IsNumeric(v) Then
        If CDbl(v) > 0 Then dt = DateSerial(1899, 12, 30) + CDbl(v) Else GoTo Fail
    ElseIf Len(Trim$(CStr(v))) > 0 Then
        dt = CDate(Trim$(CStr(v)))
    Else
        GoTo Fail
    End If
    ' Placeholder/corrupt dates (before 2000) count as missing.
    If dt < #1/1/2000# Then GoTo Fail
    TryGetExcelDate = True
    Exit Function
Fail:
    TryGetExcelDate = False
End Function


' ====================== OUTPUT ARRAY / SORTING ======================
' Builds output arrays, sorts manager rows by sales rep, ranks columns, and highlights new records.
Private Sub HighlightNewRecords(ws As Worksheet, firstRow As Long, lastRow As Long, lastDataCol As Long, _
    On Error GoTo LocalFail
    LocalFail:
        LogErr "HighlightNewRecords"
        Err.Clear
        Resume Next

        LogErr "SortManagerRowsBySalesRep"
        Err.Clear
        Resume Next

    If keepCount <= 1 Then Exit Sub
    On Error Resume Next
    For i = LBound(keepCols) To UBound(keepCols) - 1
        For j = i + 1 To UBound(keepCols)
            keyI = ManagerColumnRank(CStr(data(1, keepCols(i))), keepCols(i))
            keyJ = ManagerColumnRank(CStr(data(1, keepCols(j))), keepCols(j))
            If keyI > keyJ Then
                tmp = keepCols(i): keepCols(i) = keepCols(j): keepCols(j) = tmp
            End If
        Next j
    Next i
    On Error GoTo 0
End Sub

Private Function ManagerColumnRank(ByVal headerText As String, ByVal sourceCol As Long) As Long
    Select Case LCase$(Trim$(headerText))
        Case LCase$(REP_HEADER)
            ManagerColumnRank = 1
        Case LCase$(START_DATE_HEADER)
            ManagerColumnRank = 2
        Case LCase$(NAME_HEADER)
            ManagerColumnRank = 3
        Case LCase$("Billing Customer")
            ManagerColumnRank = 4
        Case LCase$("Prescribing Physician")
            ManagerColumnRank = 5
        Case LCase$("No Rental Equipment Provided")
            ManagerColumnRank = 6
        Case LCase$(PICKUP_DATE_HEADER)
            ManagerColumnRank = 7
        Case Else
            ManagerColumnRank = 1000 + sourceCol
    End Select
End Function

Private Function BuildManagerOutputArray(ByRef data As Variant, ByVal rowsToUse As Collection, ByVal lastCol As Long, ByVal pickupDateCol As Long) As Variant
    On Error GoTo LocalFail
    LocalFail:
        LogErr "BuildManagerOutputArray"
        Err.Clear
        Resume Next

Private Function IsExcludedManagerHeader(ByVal headerText As String) As Boolean
    On Error GoTo LocalFail
    LocalFail:
        LogErr "IsExcludedManagerHeader"
        Err.Clear
        Resume Next

    LocalFail:
        LogErr "CopyColumnFormats"
        Err.Clear
        Resume Next

        dest.Columns(c).ColumnWidth = src.Columns(c).ColumnWidth
    Next c
End Sub


Private Sub MatchSourceColumnWidthsForManager(src As Worksheet, dest As Worksheet, sourceLastCol As Long)
    On Error GoTo LocalFail
    LocalFail:
        LogErr "MatchSourceColumnWidthsForManager"
        Err.Clear
        Resume Next

        LogErr "FindHeaderColumn"
        Err.Clear
        Resume Next

    If rowNum >= 1 And rowNum <= UBound(data, 1) Then
        If keyCol > 0 And keyCol <= UBound(data, 2) Then s = NormalizeText(CStr(data(rowNum, keyCol)))
        If Len(s) = 0 And nameCol > 0 And nameCol <= UBound(data, 2) Then s = NormalizeText(CStr(data(rowNum, nameCol)))
    End If
    On Error GoTo 0
    If Len(s) = 0 Then s = "ROW-" & CStr(fallbackRow)
    GetRecordKey = s
End Function

Private Function SafeValue(v As Variant, defaultText As String) As String
    On Error GoTo LocalFail
    LocalFail:
        LogErr "SafeValue"
        Err.Clear
        Resume Next

        Resume Next

    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    If f Is Nothing Then LastUsedColumn = 0 Else LastUsedColumn = f.Column
End Function

Private Function SheetExists(sheetName As String, wb As Workbook) As Boolean
    On Error GoTo LocalFail
    LocalFail:
        LogErr "SheetExists"
        Err.Clear
        Resume Next

        LogErr "SortManagersByLastName"
        Err.Clear
        Resume Next

    p = InStrRev(fullName, " ")
    If p > 0 Then ManagerSortKey = Mid$(fullName, p + 1) & "|" & Left$(fullName, p - 1) Else ManagerSortKey = fullName
End Function

Private Sub SortTextArray(ByRef arr As Variant)
    On Error GoTo LocalFail
    LocalFail:
        LogErr "SortTextArray"
        Err.Clear
        Resume Next

        LogErr "FindPriorSourceSheet"
        Err.Clear
        Resume Next

    Dim fd As FileDialog
    Dim initialName As String
    On Error GoTo Fail
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = dialogTitle
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Workbooks", "*.xlsx;*.xlsm;*.xlsb"
        If Len(initialFolder) > 0 Then
            If Len(Dir(initialFolder, vbDirectory)) > 0 Then
                If Len(suggestedFile) > 0 And Len(Dir(suggestedFile)) > 0 Then
                    .InitialFileName = suggestedFile
                Else
                    .InitialFileName = initialFolder & Application.PathSeparator
                End If
            End If
        End If
        If .Show <> -1 Then
            SelectExcelFile = False
        Else
            SelectExcelFile = .SelectedItems(1)
        End If
    End With
    Exit Function
Fail:
    LogMsg "SelectExcelFile failed: " & Err.Description
    SelectExcelFile = False
End Function

Private Sub SaveFolderSetting(ByVal settingName As String, ByVal filePath As String)
    On Error GoTo LocalFail
    LocalFail:
        LogErr "SaveFolderSetting"
        Err.Clear
        Resume Next

        LogErr "FindMostRecentPendingPickupReport"
        Err.Clear
        Resume Next

    folder = GetSavedOutputFolder()
    If Len(folder) = 0 Or Len(Dir(folder, vbDirectory)) = 0 Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Choose where finished Pending Pickup reports should be saved"
            .AllowMultiSelect = False
            If .Show <> -1 Then Err.Raise vbObjectError + 1000, , "Save location selection was cancelled."
            folder = .SelectedItems(1)
        End With
        SaveOutputFolder folder
    End If
    baseName = "Pending Pickup " & Format(Date, "yyyy-mm-dd") & ".xlsx"
    candidate = folder & Application.PathSeparator & baseName
    If Len(Dir(CStr(candidate))) > 0 Then
        candidate = folder & Application.PathSeparator & "Pending Pickup " & Format(Now, "yyyy-mm-dd_HHMMSS") & ".xlsx"
    End If
    BuildOutputPath = CStr(candidate)
    Exit Function
Fail:
    LogMsg "BuildOutputPath failed: " & Err.Description
    Err.Raise Err.Number, , "Output folder error: " & Err.Description
End Function

Private Function GetSavedOutputFolder() As String
    On Error GoTo LocalFail
    LocalFail:
        LogErr "GetSavedOutputFolder"
        Err.Clear
        Resume Next

' report looks clean and focused only on actual data.
Private Sub HideUnusedCells(ByVal ws As Worksheet)
    On Error GoTo LocalFail
    LocalFail:
        LogErr "HideUnusedCells"
        Err.Clear
        Resume Next

