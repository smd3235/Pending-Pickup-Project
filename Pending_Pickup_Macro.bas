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
Private mRunLogPath As String
Private mInReportRun As Boolean

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
    Dim managerCol As Long, statusCol As Long, repCol As Long, nameCol As Long, keyCol As Long, pickupDateCol As Long, startDateCol As Long
    Dim priorStatusCol As Long, priorNameCol As Long, priorKeyCol As Long, priorManagerCol As Long
    Dim r As Long, i As Long
    Dim manager As String, key As Variant, status As String
    Dim ws As Worksheet, dashboard As Worksheet, rawOut As Worksheet
    Dim arrOut As Variant, arrRaw As Variant
    Dim managers As Variant
    Dim savePath As String
    Dim priorPath As Variant
    Dim oldCalc As XlCalculation, oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean
    Dim hasPrior As Boolean
    Dim currentCount As Long, priorCount As Long, newCount As Long, clearedCount As Long, stillCount As Long
    Dim newKeys As Object
    Dim age30Count As Long, age60Count As Long, age90Count As Long, missingDateCount As Long
    Dim errText As String
    Dim managerLastCol As Long, managerPickupDateCol As Long, managerStartDateCol As Long, managerAgeCol As Long, managerRepCol As Long
    Dim sortedManagerRows As Collection
    Dim managerSheetNames As Object
    Dim overallAvgAge As Variant
    Dim reportStamp As String, sourceName As String
    Dim stillByManager As Object, clearedByManager As Object

    #If Not ENABLE_ERL Then
    On Error GoTo CleanFail
    #End If
    mInReportRun = True
    mRunLogPath = BuildRunLogPath()
    LogMsg "=== CreatePendingPickupReport started ==="
    LogMsg "Run log: " & mRunLogPath
    LogMsg "Completion marker: " & COMPLETION_MARKER_PATH
    LogMsg "Excel version: " & Application.Version
    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts

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
    keyCol = FindHeaderColumn(raw, headerRow, KEY_HEADER)
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
    LogMsg "Source dimensions: rows=" & lastRow & " cols=" & lastCol & " statusCol=" & statusCol & " managerCol=" & managerCol & " keyCol=" & keyCol & " pickupDateCol=" & pickupDateCol
    Set currentRows = New Collection
    Set currentByKey = CreateObject("Scripting.Dictionary")
    currentByKey.CompareMode = vbTextCompare
    Set newKeys = CreateObject("Scripting.Dictionary")
    newKeys.CompareMode = vbTextCompare
    Set managerRows = CreateObject("Scripting.Dictionary")
    managerRows.CompareMode = vbTextCompare
    Set managerNames = CreateObject("Scripting.Dictionary")
    managerNames.CompareMode = vbTextCompare
    Set managerSheetNames = CreateObject("Scripting.Dictionary")
    managerSheetNames.CompareMode = vbTextCompare
    Set stillByManager = CreateObject("Scripting.Dictionary")
    stillByManager.CompareMode = vbTextCompare
    Set clearedByManager = CreateObject("Scripting.Dictionary")
    clearedByManager.CompareMode = vbTextCompare

    For r = 2 To UBound(currentData, 1)
        status = Trim$(CStr(currentData(r, statusCol)))
        If InStr(1, status, PENDING_TEXT, vbTextCompare) > 0 Then
            currentRows.Add r
            key = GetRecordKey(currentData, r, keyCol, nameCol, r)
            If Not currentByKey.Exists(key) Then currentByKey.Add key, r

            manager = Trim$(CStr(currentData(r, managerCol)))
            If Len(manager) = 0 Then manager = "Unassigned"
            If Not managerRows.Exists(manager) Then
                Set managerRows(manager) = New Collection
                managerNames.Add manager, manager
            End If
            managerRows(manager).Add r
        End If
    Next r

    currentCount = currentRows.Count
    LogMsg "currentCount=" & currentCount & " managers=" & managerNames.Count
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
                priorKeyCol = FindHeaderColumn(prior, priorHeaderRow, KEY_HEADER)
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
                key = GetRecordKey(priorData, r, priorKeyCol, priorNameCol, r)
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
        For Each key In currentByKey.Keys
            If priorByKey.Exists(key) Then
                stillCount = stillCount + 1
            Else
                newCount = newCount + 1
                If Not newKeys.Exists(key) Then newKeys.Add key, True
            End If
        Next key
        For Each key In priorByKey.Keys
            If Not currentByKey.Exists(key) Then clearedCount = clearedCount + 1
        Next key
    End If

    overallAvgAge = CalculateAverageAge(currentData, currentRows, pickupDateCol)

    If hasPrior Then
        For i = 1 To currentRows.Count
            r = CLng(currentRows(i))
            key = GetRecordKey(currentData, r, keyCol, nameCol, r)
            If priorByKey.Exists(key) Then
                manager = Trim$(CStr(currentData(r, managerCol)))
                If Len(manager) = 0 Then manager = "Unassigned"
                If stillByManager.Exists(manager) Then
                    stillByManager(manager) = stillByManager(manager) + 1
                Else
                    stillByManager.Add manager, 1
                End If
            End If
        Next i
        If priorManagerCol > 0 Then
            For i = 1 To priorRows.Count
                r = CLng(priorRows(i))
                key = GetRecordKey(priorData, r, priorKeyCol, priorNameCol, r)
                If Not currentByKey.Exists(key) Then
                    manager = Trim$(CStr(priorData(r, priorManagerCol)))
                    If Len(manager) = 0 Then manager = "Unassigned"
                    If clearedByManager.Exists(manager) Then
                        clearedByManager(manager) = clearedByManager(manager) + 1
                    Else
                        clearedByManager.Add manager, 1
                    End If
                End If
            Next i
        End If
    End If

    CalculateAgingStats currentData, currentRows, pickupDateCol, age30Count, age60Count, age90Count, missingDateCount
    LogMsg "Aging: 30=" & age30Count & " 60=" & age60Count & " 90=" & age90Count & " missing=" & missingDateCount

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
        managerSheetNames.Add manager, ws.Name
        Set sortedManagerRows = SortManagerRowsBySalesRep(currentData, managerRows(manager), repCol, nameCol)
        arrOut = BuildManagerOutputArray(currentData, sortedManagerRows, lastCol, pickupDateCol)
        managerLastCol = UBound(arrOut, 2)
        ws.Range(ws.Cells(MANAGER_HEADER_ROW, 1), ws.Cells(MANAGER_HEADER_ROW + UBound(arrOut, 1) - 1, managerLastCol)).Value = arrOut
        managerPickupDateCol = FindHeaderColumn(ws, MANAGER_HEADER_ROW, PICKUP_DATE_HEADER)
        managerStartDateCol = FindHeaderColumn(ws, MANAGER_HEADER_ROW, START_DATE_HEADER)
        managerAgeCol = FindHeaderColumn(ws, MANAGER_HEADER_ROW, "Pickup Age (Days)")
        managerRepCol = FindHeaderColumn(ws, MANAGER_HEADER_ROW, REP_HEADER)
        FormatManagerSheet ws, manager, managerRows(manager).Count, managerLastCol, managerPickupDateCol, managerStartDateCol, managerAgeCol, managerRepCol, currentData, sortedManagerRows, newKeys, keyCol, nameCol, hasPrior
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
                   age30Count, age60Count, age90Count, missingDateCount, hasPrior, managers, managerRows, managerSheetNames, currentData, newKeys, keyCol, nameCol, pickupDateCol, _
                   overallAvgAge, reportStamp, sourceName, stillByManager, clearedByManager, currentRows
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
    age30Count As Long, age60Count As Long, age90Count As Long, missingDateCount As Long, hasPrior As Boolean, _
    managers As Variant, managerRows As Object, managerSheetNames As Object, ByRef currentData As Variant, newKeys As Object, keyCol As Long, nameCol As Long, pickupDateCol As Long, _
    ByVal overallAvgAge As Variant, ByVal reportStamp As String, ByVal sourceName As String, ByVal stillByManager As Object, ByVal clearedByManager As Object, ByVal pendingRows As Collection)

    Dim pctChange As Variant
    Dim i As Long, rowOut As Long
    Dim managerName As String, safeName As String
    Dim mAge30 As Long, mAge60 As Long, mAge90 As Long, mMissing As Long
    Dim mNew As Variant, mAvgAge As Variant
    Dim mStill As Variant, mCleared As Variant
    Dim mPct30 As Variant
    Dim tableLastRow As Long, chartRow As Long
    Dim chartObj As ChartObject
    Dim cs As Object
    Dim chartNames() As String, chartVals() As Double
    Dim mi As Long, mj As Long, sortTmp As String, sortTmpVal As Double
    Dim agingTitleRow As Long
    Dim bucketLabels(1 To 5) As String, bucketCounts(1 To 5) As Long
    Dim agg As Variant, dt As Date, ageD As Long, srcRow As Long
    Dim k As Long

    ws.Cells.Clear
    ws.Columns("A:J").ColumnWidth = 14
    ws.Rows("1:90").RowHeight = 22

    With ws.Range("A1:H2")
        .Merge
        .Value = "PENDING PICKUP DASHBOARD"
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Name = "Segoe UI"
        .Font.Size = 22
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("A1:H2").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(246, 206, 62)
    End With
    ws.Rows("1:2").RowHeight = 30

    ' Report stamp / source line.
    With ws.Range("A3:J3")
        .Merge
        .Value = "Report generated " & reportStamp & "   |   Source: " & sourceName
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Italic = True
        .Font.Color = RGB(89, 89, 89)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Rows("3:3").RowHeight = 18

    If hasPrior And priorCount > 0 Then
        pctChange = (currentCount - priorCount) / priorCount
    Else
        pctChange = "N/A"
    End If

    ' Manager table metrics, computed once so the chart can Pareto-sort by Current.
    Dim mNewByMgr As Object, mAge30ByMgr As Object, mAge60ByMgr As Object, mAge90ByMgr As Object, mMissingByMgr As Object
    Dim mAvgByMgr As Object, mPct30ByMgr As Object, mStillByMgr As Object, mClearedByMgr As Object
    Set mNewByMgr = CreateObject("Scripting.Dictionary"): mNewByMgr.CompareMode = vbTextCompare
    Set mAge30ByMgr = CreateObject("Scripting.Dictionary"): mAge30ByMgr.CompareMode = vbTextCompare
    Set mAge60ByMgr = CreateObject("Scripting.Dictionary"): mAge60ByMgr.CompareMode = vbTextCompare
    Set mAge90ByMgr = CreateObject("Scripting.Dictionary"): mAge90ByMgr.CompareMode = vbTextCompare
    Set mMissingByMgr = CreateObject("Scripting.Dictionary"): mMissingByMgr.CompareMode = vbTextCompare
    Set mAvgByMgr = CreateObject("Scripting.Dictionary"): mAvgByMgr.CompareMode = vbTextCompare
    Set mPct30ByMgr = CreateObject("Scripting.Dictionary"): mPct30ByMgr.CompareMode = vbTextCompare
    Set mStillByMgr = CreateObject("Scripting.Dictionary"): mStillByMgr.CompareMode = vbTextCompare
    Set mClearedByMgr = CreateObject("Scripting.Dictionary"): mClearedByMgr.CompareMode = vbTextCompare

    For i = LBound(managers) To UBound(managers)
        managerName = CStr(managers(i))
        CalculateAgingStats currentData, managerRows(managerName), pickupDateCol, mAge30, mAge60, mAge90, mMissing
        mAvgAge = CalculateAverageAge(currentData, managerRows(managerName), pickupDateCol)
        If managerRows(managerName).Count > 0 Then
            mPct30 = mAge30 / CLng(managerRows(managerName).Count)
        Else
            mPct30 = 0
        End If
        mAge30ByMgr(managerName) = mAge30
        mAge60ByMgr(managerName) = mAge60
        mAge90ByMgr(managerName) = mAge90
        mMissingByMgr(managerName) = mMissing
        mAvgByMgr(managerName) = mAvgAge
        mPct30ByMgr(managerName) = mPct30
        If hasPrior Then
            mNew = CountNewManagerRows(currentData, managerRows(managerName), newKeys, keyCol, nameCol)
            If stillByManager.Exists(managerName) Then mStill = stillByManager(managerName) Else mStill = 0
            If clearedByManager.Exists(managerName) Then mCleared = clearedByManager(managerName) Else mCleared = 0
            mNewByMgr(managerName) = mNew
            mStillByMgr(managerName) = mStill
            mClearedByMgr(managerName) = mCleared
        Else
            mNewByMgr(managerName) = "N/A"
            mStillByMgr(managerName) = "N/A"
            mClearedByMgr(managerName) = "N/A"
        End If
    Next i

    ' Uniform dashboard KPI grid: all cards span equal 2-column widths (A:J)
    ' so cards align vertically with each other and with the manager table below.
    AddDashboardCard ws, "A4:B6", "CURRENT", "Pending Pickups", currentCount, RGB(255, 242, 204)
    AddDashboardCard ws, "C4:D6", "PRIOR", "Pending Pickups", IIf(hasPrior, priorCount, "N/A"), RGB(221, 235, 247)
    AddDashboardCard ws, "E4:F6", "NEW", "Since Prior Report", IIf(hasPrior, newCount, "N/A"), RGB(226, 239, 218)
    AddDashboardCard ws, "G4:H6", "CLEARED", "Since Prior Report", IIf(hasPrior, clearedCount, "N/A"), RGB(244, 204, 204)
    AddDashboardCard ws, "I4:J6", "STILL PENDING", "In Both Reports", IIf(hasPrior, stillCount, "N/A"), RGB(217, 234, 247)

    AddDashboardCard ws, "A8:B10", "% CHANGE", "From Prior Report", pctChange, RGB(234, 209, 220)
    If hasPrior And priorCount > 0 Then
        ws.Range("A8:B10").NumberFormat = "0.0%"
        ws.Cells(10, 1).Value = pctChange
    End If
    AddDashboardCard ws, "C8:D10", "30+ DAYS", "Older Pending Pickups", age30Count, RGB(244, 204, 204)
    AddDashboardCard ws, "E8:F10", "60+ DAYS", "Older Pending Pickups", age60Count, RGB(234, 153, 153)
    AddDashboardCard ws, "G8:H10", "90+ DAYS", "Older Pending Pickups", age90Count, RGB(224, 102, 102)
    AddDashboardCard ws, "I8:J10", "MISSING DATE", "AUDIT: TP End Date", missingDateCount, RGB(255, 242, 204)

    ' Manager breakdown: clickable manager names plus useful manager-level metrics.
    ws.Range("A12:K12").Merge
    ws.Range("A12").Value = "PENDING PICKUPS BY MANAGER"
    StyleSectionHeader ws.Range("A12:K12")
    ws.Cells(13, 1).Value = "Manager"
    ws.Cells(13, 2).Value = "Current"
    ws.Cells(13, 3).Value = "New"
    ws.Cells(13, 4).Value = "30+ Days"
    ws.Cells(13, 5).Value = "60+ Days"
    ws.Cells(13, 6).Value = "90+ Days"
    ws.Cells(13, 7).Value = "Missing Date"
    ws.Cells(13, 8).Value = "Avg Age"
    ws.Cells(13, 9).Value = "% 30+"
    ws.Cells(13, 10).Value = "Cleared"
    ws.Cells(13, 11).Value = "Still Pending"
    StyleTableHeader ws.Range("A13:K13")

    rowOut = 14
    For i = LBound(managers) To UBound(managers)
        managerName = CStr(managers(i))
        safeName = CStr(managerSheetNames(managerName))
        ws.Hyperlinks.Add Anchor:=ws.Cells(rowOut, 1), Address:="", SubAddress:="'" & Replace(safeName, "'", "''") & "'!A1", TextToDisplay:=managerName
        ws.Cells(rowOut, 2).Value = managerRows(managerName).Count
        ws.Cells(rowOut, 3).Value = mNewByMgr(managerName)
        ws.Cells(rowOut, 4).Value = mAge30ByMgr(managerName)
        ws.Cells(rowOut, 5).Value = mAge60ByMgr(managerName)
        ws.Cells(rowOut, 6).Value = mAge90ByMgr(managerName)
        ws.Cells(rowOut, 7).Value = mMissingByMgr(managerName)
        ws.Cells(rowOut, 8).Value = mAvgByMgr(managerName)
        ws.Cells(rowOut, 8).NumberFormat = "0.0"
        ws.Cells(rowOut, 9).Value = mPct30ByMgr(managerName)
        ws.Cells(rowOut, 9).NumberFormat = "0.0%"
        ws.Cells(rowOut, 10).Value = mClearedByMgr(managerName)
        ws.Cells(rowOut, 11).Value = mStillByMgr(managerName)
        rowOut = rowOut + 1
    Next i

    tableLastRow = rowOut - 1
    If rowOut > 14 Then
        StyleDashboardTable ws.Range(ws.Cells(14, 1), ws.Cells(tableLastRow, 11))
        ws.Range(ws.Cells(14, 8), ws.Cells(tableLastRow, 8)).NumberFormat = "0.0"
        ws.Range(ws.Cells(14, 9), ws.Cells(tableLastRow, 9)).NumberFormat = "0.0%"
        ws.Range(ws.Cells(14, 1), ws.Cells(tableLastRow, 1)).HorizontalAlignment = xlLeft
        ws.Range(ws.Cells(14, 1), ws.Cells(tableLastRow, 1)).Font.Underline = xlUnderlineStyleSingle

        ' Visual cues: data bars on Current, colour scale on 90+ Days and % 30+.
        On Error Resume Next
        ws.Range(ws.Cells(14, 2), ws.Cells(tableLastRow, 2)).FormatConditions.AddDatabar
        Set cs = ws.Range(ws.Cells(14, 6), ws.Cells(tableLastRow, 6)).FormatConditions.AddColorScale(3)
        If Not cs Is Nothing Then
            cs.ColorScaleCriteria(1).FormatColor.Color = RGB(255, 255, 255)
            cs.ColorScaleCriteria(2).FormatColor.Color = RGB(255, 192, 0)
            cs.ColorScaleCriteria(3).FormatColor.Color = RGB(192, 0, 0)
        End If
        Set cs = ws.Range(ws.Cells(14, 9), ws.Cells(tableLastRow, 9)).FormatConditions.AddColorScale(3)
        If Not cs Is Nothing Then
            cs.ColorScaleCriteria(1).FormatColor.Color = RGB(198, 239, 206)
            cs.ColorScaleCriteria(2).FormatColor.Color = RGB(255, 235, 156)
            cs.ColorScaleCriteria(3).FormatColor.Color = RGB(255, 199, 206)
        End If
        On Error GoTo 0

        ' Totals row.
        ws.Cells(rowOut, 1).Value = "TOTAL"
        ws.Cells(rowOut, 2).Value = currentCount
        ws.Cells(rowOut, 3).Value = IIf(hasPrior, newCount, "N/A")
        ws.Cells(rowOut, 4).Value = age30Count
        ws.Cells(rowOut, 5).Value = age60Count
        ws.Cells(rowOut, 6).Value = age90Count
        ws.Cells(rowOut, 7).Value = missingDateCount
        ws.Cells(rowOut, 8).Value = overallAvgAge
        ws.Cells(rowOut, 8).NumberFormat = "0.0"
        ws.Cells(rowOut, 9).Value = IIf(currentCount > 0, age30Count / currentCount, 0)
        ws.Cells(rowOut, 9).NumberFormat = "0.0%"
        ws.Cells(rowOut, 10).Value = IIf(hasPrior, clearedCount, "N/A")
        ws.Cells(rowOut, 11).Value = IIf(hasPrior, stillCount, "N/A")
        With ws.Range(ws.Cells(rowOut, 1), ws.Cells(rowOut, 11))
            .Font.Bold = True
            .Interior.Color = RGB(217, 234, 247)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlHairline
            .HorizontalAlignment = xlCenter
        End With
        ws.Cells(rowOut, 1).HorizontalAlignment = xlLeft
        ws.Cells(rowOut, 10).NumberFormat = "0"
        ws.Cells(rowOut, 11).NumberFormat = "0"
    End If

    ' Keep every dashboard section on the same clean 11-column grid.
    ws.Columns("A:K").VerticalAlignment = xlCenter
    ws.Columns("A:K").HorizontalAlignment = xlCenter
    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B:K").ColumnWidth = 13
    ws.Rows("1:2").RowHeight = 30
    ws.Rows("3:3").RowHeight = 18
    ws.Rows("4:10").RowHeight = 30
    ws.Rows("12:12").RowHeight = 24
    ws.Rows("13:13").RowHeight = 30
    If rowOut > 14 Then ws.Rows("14:" & rowOut).RowHeight = 22

    ' Bar chart of current pending pickups by manager, Pareto-sorted (largest first), placed below the table.
    If rowOut > 14 Then
        ReDim chartNames(1 To tableLastRow - 13)
        ReDim chartVals(1 To tableLastRow - 13)
        For i = 1 To tableLastRow - 13
            chartNames(i) = CStr(ws.Cells(13 + i, 1).Value)
            If IsNumeric(ws.Cells(13 + i, 2).Value) Then
                chartVals(i) = CDbl(ws.Cells(13 + i, 2).Value)
            Else
                chartVals(i) = 0
            End If
        Next i
        For mi = 1 To UBound(chartNames) - 1
            For mj = mi + 1 To UBound(chartNames)
                If chartVals(mj) > chartVals(mi) Then
                    sortTmp = chartNames(mi): chartNames(mi) = chartNames(mj): chartNames(mj) = sortTmp
                    sortTmpVal = chartVals(mi): chartVals(mi) = chartVals(mj): chartVals(mj) = sortTmpVal
                End If
            Next mj
        Next mi

        chartRow = rowOut + 2
        ws.Range(ws.Cells(chartRow, 1), ws.Cells(chartRow, 11)).Merge
        ws.Cells(chartRow, 1).Value = "PENDING PICKUPS BY MANAGER (CHART)"
        StyleSectionHeader ws.Range(ws.Cells(chartRow, 1), ws.Cells(chartRow, 11))
        Set chartObj = ws.ChartObjects.Add(ws.Cells(chartRow + 1, 1).Left, ws.Cells(chartRow + 1, 1).Top, 900, 330)
        With chartObj.Chart
            .ChartType = xlColumnClustered
            .SeriesCollection.NewSeries
            .SeriesCollection(1).Values = chartVals
            .SeriesCollection(1).XValues = chartNames
            .SeriesCollection(1).Name = "Pending Pickups"
            .HasTitle = True
            .ChartTitle.Text = "Current Pending Pickups by Manager"
            .HasLegend = False
            .Axes(xlCategory).TickLabels.Font.Size = 9
            .Axes(xlValue).TickLabels.Font.Size = 9
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.Text = "Pending Pickups"
            .Axes(xlValue).AxisTitle.Font.Size = 9
            .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(31, 78, 121)
        End With
    End If

    ' Aging breakdown: buckets derived from the aging KPIs computed earlier.
    If chartRow = 0 Then chartRow = rowOut + 2
    bucketLabels(1) = "0-30 Days": bucketCounts(1) = currentCount - age30Count
    bucketLabels(2) = "31-60 Days": bucketCounts(2) = age30Count - age60Count
    bucketLabels(3) = "61-90 Days": bucketCounts(3) = age60Count - age90Count
    bucketLabels(4) = "90+ Days": bucketCounts(4) = age90Count
    bucketLabels(5) = "Missing Date": bucketCounts(5) = missingDateCount

    agingTitleRow = chartRow + 20
    ws.Range(ws.Cells(agingTitleRow, 1), ws.Cells(agingTitleRow, 11)).Merge
    ws.Cells(agingTitleRow, 1).Value = "AGING BREAKDOWN"
    StyleSectionHeader ws.Range(ws.Cells(agingTitleRow, 1), ws.Cells(agingTitleRow, 11))
    ws.Cells(agingTitleRow + 1, 1).Value = "Age Bucket"
    ws.Cells(agingTitleRow + 1, 2).Value = "Records"
    ws.Cells(agingTitleRow + 1, 3).Value = "% of Total"
    StyleTableHeader ws.Range(ws.Cells(agingTitleRow + 1, 1), ws.Cells(agingTitleRow + 1, 3))
    For k = 1 To 5
        ws.Cells(agingTitleRow + 1 + k, 1).Value = bucketLabels(k)
        ws.Cells(agingTitleRow + 1 + k, 2).Value = bucketCounts(k)
        ws.Cells(agingTitleRow + 1 + k, 3).Value = IIf(currentCount > 0, bucketCounts(k) / currentCount, 0)
        ws.Cells(agingTitleRow + 1 + k, 3).NumberFormat = "0.0%"
    Next k
    StyleDashboardTable ws.Range(ws.Cells(agingTitleRow + 1, 1), ws.Cells(agingTitleRow + 6, 3))
    Set chartObj = ws.ChartObjects.Add(ws.Cells(agingTitleRow + 1, 5).Left, ws.Cells(agingTitleRow + 1, 1).Top, 560, 240)
    With chartObj.Chart
        .ChartType = xlColumnClustered
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Values = ws.Range(ws.Cells(agingTitleRow + 2, 2), ws.Cells(agingTitleRow + 6, 2))
        .SeriesCollection(1).XValues = ws.Range(ws.Cells(agingTitleRow + 2, 1), ws.Cells(agingTitleRow + 6, 1))
        .SeriesCollection(1).Name = "Pending Pickups"
        .HasTitle = True
        .ChartTitle.Text = "Pending Pickup Age Distribution"
        .HasLegend = False
        .Axes(xlCategory).TickLabels.Font.Size = 9
        .Axes(xlValue).TickLabels.Font.Size = 9
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(31, 78, 121)
    End With


    ws.Tab.Color = RGB(31, 78, 121)
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With

    ws.Activate
    AddLogoFromTemplate ws
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 90
    On Error GoTo 0
End Sub

Private Sub AddLogoFromTemplate(ws As Worksheet)
    Dim src As Shape, shp As Shape
    On Error Resume Next
    Set src = ThisWorkbook.Worksheets("START HERE").Shapes("AOTILogo")
    If src Is Nothing Then Exit Sub

    ws.Range("A3").Select
    src.Copy
    DoEvents
    ws.Paste
    Set shp = ws.Shapes(ws.Shapes.Count)
    If shp Is Nothing Then Exit Sub
    shp.LockAspectRatio = msoTrue
    shp.Width = ws.Range("I1:J2").Width - 10
    shp.Left = ws.Range("I1").Left + 5
    shp.Top = ws.Range("I1").Top + 5
End Sub

Private Sub AddDashboardCard(ws As Worksheet, address As String, labelText As String, subText As String, valueText As Variant, fillColor As Long)
    Dim rg As Range, labelRg As Range, valueRow As Range
    Set rg = ws.Range(address)
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
        .Font.Size = 9
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    valueRow.Merge
    valueRow.Value = CStr(valueText)
    With valueRow
        .Font.Bold = True
        .Font.Name = "Segoe UI"
        .Font.Size = 20
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

' ====================== MANAGER SHEET FORMATTING ======================
' Formats one manager tab: KPI cards, data table, banding, highlighting, age warnings.
Private Sub FormatManagerSheet(ws As Worksheet, managerName As String, pendingCount As Long, managerLastCol As Long, managerPickupDateCol As Long, managerStartDateCol As Long, managerAgeCol As Long, managerRepCol As Long, _
    ByRef sourceData As Variant, rowsToUse As Collection, newKeys As Object, keyCol As Long, nameCol As Long, hasPrior As Boolean)
    Dim age30 As Long, age60 As Long, age90 As Long, missingDate As Long
    Dim managerNewCount As Variant
    Dim lastDataRow As Long
    Dim pickupCol As Long
    pickupCol = FindHeaderColumnFromArray(sourceData, PICKUP_DATE_HEADER)
    LogMsg "FormatManagerSheet start: " & managerName & " pickupCol=" & pickupCol
    CalculateAgingStats sourceData, rowsToUse, pickupCol, age30, age60, age90, missingDate
    Dim managerAvgAge As Variant
    managerAvgAge = CalculateAverageAge(sourceData, rowsToUse, pickupCol)
    Dim managerPct30 As Variant
    If rowsToUse.Count > 0 Then managerPct30 = age30 / CLng(rowsToUse.Count) Else managerPct30 = 0

    If hasPrior Then
        managerNewCount = CountNewManagerRows(sourceData, rowsToUse, newKeys, keyCol, nameCol)
    Else
        managerNewCount = "N/A"
    End If
    LogMsg "FormatManagerSheet: " & managerName & " done stats"

    With ws.Range(ws.Cells(1, 1), ws.Cells(2, managerLastCol))
        .Merge
        .Value = managerName & "   |   as of " & Format(Date, "mm/dd/yyyy")
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Name = "Segoe UI"
        .Font.Size = 18
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(1, 1), ws.Cells(2, managerLastCol)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(246, 206, 62)
    End With
    ws.Rows("1:2").RowHeight = 28

    ' KPI cards, fully contained inside the manager data width.
    AddDashboardCard ws, "A4:B6", "PENDING PICKUPS", "Current", pendingCount, RGB(255, 242, 204)
    AddDashboardCard ws, "C4:D6", "SALES REPS", "Total Reps", CountUniqueManagerRepsOnSheet(ws, MANAGER_HEADER_ROW + 1, LastUsedRow(ws), managerRepCol), RGB(221, 235, 247)
    AddDashboardCard ws, "E4:F6", "NEW", "Since Prior Report", managerNewCount, RGB(226, 239, 218)
    AddDashboardCard ws, "G4:H6", "30+ DAYS", "Older Pending Pickups", age30, RGB(244, 204, 204)
    AddDashboardCard ws, "A8:B10", "60+ DAYS", "Older Pending Pickups", age60, RGB(234, 153, 153)
    AddDashboardCard ws, "C8:D10", "90+ DAYS", "Older Pending Pickups", age90, RGB(224, 102, 102)
    AddDashboardCard ws, "E8:F10", "MISSING DATE", "AUDIT: TP End Date", missingDate, RGB(255, 242, 204)
    AddDashboardCard ws, "G8:H10", "AVG AGE", "Pending Pickup Age", IIf(IsNumeric(managerAvgAge), Format(managerAvgAge, "0.0"), "N/A"), RGB(209, 226, 242)
    AddDashboardCard ws, "A13:B14", "% 30+", "Older Pending Pickups", IIf(rowsToUse.Count > 0, managerPct30, 0), RGB(234, 209, 220)
    If rowsToUse.Count > 0 Then ws.Range("A13:B14").NumberFormat = "0.0%"

    lastDataRow = LastUsedRow(ws)
    LogMsg "  lastDataRow=" & lastDataRow
    If lastDataRow >= MANAGER_HEADER_ROW Then
        LogMsg "  ApplyDataFormatting"
        ApplyDataFormatting ws, MANAGER_HEADER_ROW, lastDataRow, managerLastCol
        LogMsg "  ApplyDateFormatsToOutput"
        ApplyDateFormatsToOutput ws, MANAGER_HEADER_ROW, lastDataRow, managerStartDateCol, managerPickupDateCol
        LogMsg "  ApplyAgeWarningFont"
        If managerAgeCol > 0 And lastDataRow >= MANAGER_HEADER_ROW + 1 Then
            ws.Range(ws.Cells(MANAGER_HEADER_ROW + 1, managerAgeCol), ws.Cells(lastDataRow, managerAgeCol)).NumberFormat = "0"
        End If
        LogMsg "  ApplyRepBanding"
        ApplyRepBanding ws, MANAGER_HEADER_ROW + 1, lastDataRow, managerLastCol, managerRepCol
        If hasPrior Then HighlightNewRecords ws, MANAGER_HEADER_ROW + 1, lastDataRow, managerLastCol, sourceData, rowsToUse, newKeys, keyCol, nameCol
        If managerAgeCol > 0 Then ApplyAgeWarningFont ws, MANAGER_HEADER_ROW + 1, lastDataRow, managerAgeCol
        LogMsg "  HighlightMissingPickupDates"
        HighlightMissingPickupDates ws, MANAGER_HEADER_ROW + 1, lastDataRow, managerPickupDateCol
        LogMsg "  FormatYesNoColumns"
        FormatYesNoColumns ws, MANAGER_HEADER_ROW, lastDataRow, managerLastCol
        LogMsg "  formatting done"
    End If

    ' The data table is the exact same width as the KPI/header area.
    SetManagerColumnWidths ws, managerLastCol
    ws.Rows("4:6").RowHeight = 26
    ws.Rows("8:10").RowHeight = 26
    ws.Rows(CStr(MANAGER_HEADER_ROW) & ":" & CStr(MANAGER_HEADER_ROW)).RowHeight = 34

    ws.Tab.Color = RGB(189, 215, 238)
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = CStr(MANAGER_HEADER_ROW) & ":" & CStr(MANAGER_HEADER_ROW)
    End With

    ws.Activate
    On Error Resume Next
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitRow = MANAGER_HEADER_ROW
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 85
    On Error GoTo 0
End Sub

' ====================== TABLE / ROW FORMATTING ======================
' Cosmetic helpers: rep banding, age warning font, missing-date highlight, Yes/No formatting, column widths, raw-data formatting, style primitives.
Private Sub ApplyRepBanding(ws As Worksheet, firstRow As Long, lastRow As Long, lastDataCol As Long, repCol As Long)
    Dim r As Long, repName As String, priorRep As String
    Dim useAlt As Boolean
    If firstRow > lastRow Or repCol <= 0 Then Exit Sub
    On Error Resume Next
    priorRep = ""
    useAlt = False
    For r = firstRow To lastRow
        repName = NormalizeText(CStr(ws.Cells(r, repCol).Value))
        If r = firstRow Then
            useAlt = False
        ElseIf repName <> priorRep Then
            useAlt = Not useAlt
        End If
        If useAlt Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, lastDataCol)).Interior.Color = RGB(217, 225, 242)
        Else
            ws.Range(ws.Cells(r, 1), ws.Cells(r, lastDataCol)).Interior.Color = RGB(255, 255, 255)
        End If
        priorRep = repName
    Next r
    On Error GoTo 0
End Sub

Private Sub ApplyAgeWarningFont(ws As Worksheet, firstRow As Long, lastRow As Long, ageCol As Long)
    Dim r As Long, v As Variant
    If ageCol <= 0 Or firstRow > lastRow Then Exit Sub
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
        minWidth = 12
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
    Dim r As Long
    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, lastCol))
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlHairline
    End With
    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, lastCol))
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .AutoFilter
    End With
    ws.Rows(headerRow).RowHeight = 34
    For r = headerRow + 1 To lastRow
        If (r - headerRow) Mod 2 = 0 Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = RGB(242, 242, 242)
        Else
            ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = RGB(255, 255, 255)
        End If
        ws.Rows(r).RowHeight = 22
    Next r
End Sub

Private Sub StyleSectionHeader(rg As Range)
    rg.Interior.Color = RGB(31, 78, 121)
    rg.Font.Color = RGB(255, 255, 255)
    rg.Font.Bold = True
    rg.Font.Name = "Segoe UI"
    rg.HorizontalAlignment = xlCenter
    rg.VerticalAlignment = xlCenter
    With rg.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(246, 206, 62)
    End With
End Sub

Private Sub StyleTableHeader(rg As Range)
    rg.Interior.Color = RGB(217, 234, 247)
    rg.Font.Bold = True
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

Private Function CalculateAverageAge(ByRef data As Variant, ByVal rowsToUse As Collection, ByVal pickupDateCol As Long) As Variant
    Dim i As Long, srcRow As Long, ageDays As Long, totalAge As Double, countAge As Long, dt As Date
    If pickupDateCol <= 0 Then
        CalculateAverageAge = "N/A"
        Exit Function
    End If
    On Error Resume Next
    For i = 1 To rowsToUse.Count
        srcRow = CLng(rowsToUse(i))
        If srcRow >= 1 And srcRow <= UBound(data, 1) Then
            If TryGetExcelDate(data(srcRow, pickupDateCol), dt) Then
                ageDays = DateDiff("d", DateValue(dt), Date)
                If ageDays < 0 Then ageDays = 0
                totalAge = totalAge + ageDays
                countAge = countAge + 1
            End If
        End If
    Next i
    On Error GoTo 0
    If countAge = 0 Then
        CalculateAverageAge = "N/A"
    Else
        CalculateAverageAge = totalAge / countAge
    End If
End Function

Private Function CountNewManagerRows(ByRef data As Variant, ByVal rowsToUse As Collection, ByVal newKeys As Object, ByVal keyCol As Long, ByVal nameCol As Long) As Long
    Dim i As Long, srcRow As Long, key As String, countNew As Long
    If newKeys Is Nothing Then Exit Function
    If newKeys.Count = 0 Then Exit Function
    On Error Resume Next
    For i = 1 To rowsToUse.Count
        srcRow = CLng(rowsToUse(i))
        If srcRow >= 1 And srcRow <= UBound(data, 1) Then
            key = GetRecordKey(data, srcRow, keyCol, nameCol, srcRow)
            If newKeys.Exists(key) Then countNew = countNew + 1
        End If
    Next i
    On Error GoTo 0
    CountNewManagerRows = countNew
End Function

Private Sub CalculateAgingStats(ByRef data As Variant, ByVal rowsToUse As Collection, ByVal pickupDateCol As Long, _
    ByRef age30 As Long, ByRef age60 As Long, ByRef age90 As Long, ByRef missingDate As Long)
    Dim i As Long, srcRow As Long, ageDays As Long, dt As Date
    age30 = 0: age60 = 0: age90 = 0: missingDate = 0
    If pickupDateCol <= 0 Then
        missingDate = rowsToUse.Count
        Exit Sub
    End If
    On Error Resume Next
    For i = 1 To rowsToUse.Count
        srcRow = CLng(rowsToUse(i))
        If srcRow >= 1 And srcRow <= UBound(data, 1) Then
            If TryGetExcelDate(data(srcRow, pickupDateCol), dt) Then
                ageDays = DateDiff("d", DateValue(dt), Date)
                If ageDays < 0 Then ageDays = 0
                If ageDays > 30 Then age30 = age30 + 1
                If ageDays > 60 Then age60 = age60 + 1
                If ageDays > 90 Then age90 = age90 + 1
            Else
                missingDate = missingDate + 1
            End If
        Else
            missingDate = missingDate + 1
        End If
    Next i
    On Error GoTo 0
End Sub

' ====================== DATE / TIME HELPERS ======================
' Date normalisation and formatting across source, output, and per-sheet ranges.
Private Sub NormalizeDateColumn(ws As Worksheet, headerRow As Long, lastRow As Long, dateCol As Long)
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
    If lastRow <= headerRow Then Exit Sub
    If startDateCol > 0 Then ws.Range(ws.Cells(headerRow + 1, startDateCol), ws.Cells(lastRow, startDateCol)).NumberFormat = "mm/dd/yyyy"
    If pickupDateCol > 0 Then ws.Range(ws.Cells(headerRow + 1, pickupDateCol), ws.Cells(lastRow, pickupDateCol)).NumberFormat = "mm/dd/yyyy"
End Sub

Private Sub ApplyDateFormatByHeader(ws As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long, ByVal headerText As String)
    Dim col As Long
    col = FindHeaderColumn(ws, headerRow, headerText)
    If col > 0 And lastRow > headerRow Then
        ws.Range(ws.Cells(headerRow + 1, col), ws.Cells(lastRow, col)).NumberFormat = "mm/dd/yyyy"
    End If
End Sub

Private Function TryGetExcelDate(ByVal v As Variant, ByRef dt As Date) As Boolean
    On Error GoTo Fail
    If IsDate(v) Then
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
    ByRef sourceData As Variant, rowsToUse As Collection, newKeys As Object, keyCol As Long, nameCol As Long)
    Dim r As Long, i As Long, key As String
    If newKeys Is Nothing Then Exit Sub
    If newKeys.Count = 0 Then Exit Sub
    On Error Resume Next
    For r = firstRow To lastRow
        i = r - firstRow + 1
        If i <= rowsToUse.Count Then
            key = GetRecordKey(sourceData, CLng(rowsToUse(i)), keyCol, nameCol, CLng(rowsToUse(i)))
            If newKeys.Exists(key) Then
                With ws.Range(ws.Cells(r, 1), ws.Cells(r, lastDataCol))
                    .Interior.Color = RGB(226, 239, 218)
                    .Font.Bold = True
                End With
            End If
        End If
    Next r
    On Error GoTo 0
End Sub

Private Function BuildOutputArray(data As Variant, rowsToUse As Collection, lastCol As Long) As Variant
    Dim arr() As Variant, i As Long, c As Long, srcRow As Long
    On Error Resume Next
    ReDim arr(1 To rowsToUse.Count + 1, 1 To lastCol)
    For c = 1 To lastCol: arr(1, c) = data(1, c): Next c
    For i = 1 To rowsToUse.Count
        srcRow = CLng(rowsToUse(i))
        If srcRow >= 1 And srcRow <= UBound(data, 1) Then
            For c = 1 To lastCol: arr(i + 1, c) = data(srcRow, c): Next c
        End If
    Next i
    On Error GoTo 0
    BuildOutputArray = arr
End Function

Private Function SortManagerRowsBySalesRep(ByRef data As Variant, ByVal rowsToUse As Collection, ByVal repCol As Long, ByVal nameCol As Long) As Collection
    Dim arr() As Long, i As Long, j As Long, tmp As Long
    Dim keyI As String, keyJ As String
    Dim result As New Collection
    If rowsToUse.Count = 0 Then
        Set SortManagerRowsBySalesRep = result
        Exit Function
    End If
    On Error Resume Next
    ReDim arr(1 To rowsToUse.Count)
    For i = 1 To rowsToUse.Count
        arr(i) = CLng(rowsToUse(i))
    Next i
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) >= 1 And arr(i) <= UBound(data, 1) And arr(j) >= 1 And arr(j) <= UBound(data, 1) Then
                keyI = NormalizeText(CStr(data(arr(i), repCol))) & "|" & NormalizeText(CStr(data(arr(i), nameCol)))
                keyJ = NormalizeText(CStr(data(arr(j), repCol))) & "|" & NormalizeText(CStr(data(arr(j), nameCol)))
                If StrComp(keyI, keyJ, vbTextCompare) > 0 Then
                    tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
                End If
            End If
        Next j
    Next i
    For i = LBound(arr) To UBound(arr)
        result.Add arr(i)
    Next i
    On Error GoTo 0
    Set SortManagerRowsBySalesRep = result
End Function

Private Sub SortManagerColumns(ByRef keepCols() As Long, ByVal keepCount As Long, ByRef data As Variant)
    Dim i As Long, j As Long, tmp As Long
    Dim keyI As Long, keyJ As Long
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
    Dim keepCols() As Long, keepCount As Long
    Dim arr() As Variant, i As Long, c As Long, srcRow As Long
    Dim headerText As String, pickupOutCol As Long, ageRef As String
    On Error Resume Next
    keepCount = 0
    For c = 1 To lastCol
        headerText = Trim$(CStr(data(1, c)))
        If Not IsExcludedManagerHeader(headerText) Then
            keepCount = keepCount + 1
            ReDim Preserve keepCols(1 To keepCount)
            keepCols(keepCount) = c
        End If
    Next c

    ' Manager-tab order: Sales Representative, Start Date, Name, then the remaining retained source columns.
    SortManagerColumns keepCols, keepCount, data

    ' Add the calculated age as the final manager-tab column.
    ReDim arr(1 To rowsToUse.Count + 1, 1 To keepCount + 1)
    pickupOutCol = 0
    For c = 1 To keepCount
        arr(1, c) = data(1, keepCols(c))
        If StrComp(Trim$(CStr(data(1, keepCols(c)))), PICKUP_DATE_HEADER, vbTextCompare) = 0 Then pickupOutCol = c
    Next c
    arr(1, keepCount + 1) = "Pickup Age (Days)"

    ' Age is a live formula so the report ages without regenerating; blank or
    ' placeholder dates (before 2000) show as blank.
    For i = 1 To rowsToUse.Count
        srcRow = CLng(rowsToUse(i))
        If srcRow >= 1 And srcRow <= UBound(data, 1) Then
            For c = 1 To keepCount
                arr(i + 1, c) = data(srcRow, keepCols(c))
            Next c
        End If
        If pickupOutCol > 0 Then
            ageRef = ColumnLetter(pickupOutCol) & (MANAGER_HEADER_ROW + i)
            arr(i + 1, keepCount + 1) = "=IF(AND(ISNUMBER(" & ageRef & "), " & ageRef & ">=DATE(2000,1,1)), TODAY() - " & ageRef & ", """")"
        Else
            arr(i + 1, keepCount + 1) = ""
        End If
    Next i
    On Error GoTo 0
    BuildManagerOutputArray = arr
End Function

Private Function ColumnLetter(ByVal col As Long) As String
    Dim s As String, n As Long
    n = col
    Do While n > 0
        s = Chr$(65 + (n - 1) Mod 26) & s
        n = (n - 1) \ 26
    Loop
    ColumnLetter = s
End Function


' ====================== COLUMN / HEADER UTILITIES ======================
' Header exclusion rules, header lookup from arrays, and column width matching between source and output sheets.
Private Function IsExcludedManagerHeader(ByVal headerText As String) As Boolean
    Select Case LCase$(Trim$(headerText))
        Case LCase$("EMPI"), LCase$("Customer"), LCase$("Treatment Status"), LCase$("Consumable Type"), _
             LCase$("Assigned Setup Technician"), LCase$("Admin Representative"), LCase$("Patient Date of Birth"), _
             LCase$("Internal ID"), LCase$("Last Modified"), LCase$("Last Name"), LCase$("Assigned Sales Manager"), LCase$("Assigned Associate WCC")
            IsExcludedManagerHeader = True
        Case Else
            IsExcludedManagerHeader = False
    End Select
End Function

Private Function FindHeaderColumnFromArray(ByRef data As Variant, ByVal headerText As String) As Long
    Dim c As Long
    On Error Resume Next
    For c = 1 To UBound(data, 2)
        If StrComp(Trim$(CStr(data(1, c))), headerText, vbTextCompare) = 0 Then
            FindHeaderColumnFromArray = c
            Exit Function
        End If
    Next c
    On Error GoTo 0
End Function


Private Sub CopyColumnFormats(src As Worksheet, dest As Worksheet, lastCol As Long)
    Dim c As Long
    On Error Resume Next
    For c = 1 To lastCol
        dest.Columns(c).ColumnWidth = src.Columns(c).ColumnWidth
        dest.Columns(c).NumberFormat = src.Columns(c).NumberFormat
    Next c
    On Error GoTo 0
End Sub

Private Sub MatchSourceColumnWidths(src As Worksheet, dest As Worksheet, lastCol As Long)
    Dim c As Long
    For c = 1 To lastCol
        dest.Columns(c).ColumnWidth = src.Columns(c).ColumnWidth
    Next c
End Sub


Private Sub MatchSourceColumnWidthsForManager(src As Worksheet, dest As Worksheet, sourceLastCol As Long)
    Dim c As Long, srcCol As Long, headerText As String
    For c = 1 To LastUsedColumn(dest)
        headerText = Trim$(CStr(dest.Cells(MANAGER_HEADER_ROW, c).Value))
        srcCol = FindHeaderColumn(src, FindHeaderRow(src), headerText)
        If srcCol > 0 Then dest.Columns(c).ColumnWidth = src.Columns(srcCol).ColumnWidth
    Next c
End Sub

' ====================== SHEET / WORKBOOK UTILITIES ======================
' Sheet bounds, header detection, record keys, text normalisation, and sheet-name sanitation.
Private Function FindHeaderRow(ws As Worksheet) As Long
    Dim r As Long, maxRow As Long
    maxRow = WorksheetFunction.Min(25, LastUsedRow(ws))
    For r = 1 To maxRow
        If FindHeaderColumn(ws, r, MANAGER_HEADER) > 0 And _
           FindHeaderColumn(ws, r, STATUS_HEADER) > 0 And _
           FindHeaderColumn(ws, r, REP_HEADER) > 0 And _
           FindHeaderColumn(ws, r, NAME_HEADER) > 0 Then
            FindHeaderRow = r
            Exit Function
        End If
    Next r
End Function

Private Function FindHeaderColumn(ws As Worksheet, headerRow As Long, headerText As String) As Long
    Dim c As Long, lastCol As Long
    lastCol = LastUsedColumn(ws)
    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(headerRow, c).Value)), headerText, vbTextCompare) = 0 Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
End Function

Private Function GetRecordKey(ByRef data As Variant, ByVal rowNum As Long, ByVal keyCol As Long, ByVal nameCol As Long, ByVal fallbackRow As Long) As String
    Dim s As String
    On Error Resume Next
    If rowNum >= 1 And rowNum <= UBound(data, 1) Then
        If keyCol > 0 And keyCol <= UBound(data, 2) Then s = NormalizeText(CStr(data(rowNum, keyCol)))
        If Len(s) = 0 And nameCol > 0 And nameCol <= UBound(data, 2) Then s = NormalizeText(CStr(data(rowNum, nameCol)))
    End If
    On Error GoTo 0
    If Len(s) = 0 Then s = "ROW-" & CStr(fallbackRow)
    GetRecordKey = s
End Function

Private Function SafeValue(v As Variant, defaultText As String) As String
    Dim s As String
    s = Trim$(CStr(v))
    If Len(s) = 0 Then SafeValue = defaultText Else SafeValue = s
End Function

Private Function NormalizeText(s As String) As String
    NormalizeText = LCase$(Application.WorksheetFunction.Trim(Trim$(s)))
End Function

Private Function LastUsedRow(ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    If f Is Nothing Then LastUsedRow = 0 Else LastUsedRow = f.Row
End Function

Private Function LastUsedColumn(ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    If f Is Nothing Then LastUsedColumn = 0 Else LastUsedColumn = f.Column
End Function

Private Function SheetExists(sheetName As String, wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function SafeSheetName(ByVal s As String, wb As Workbook) As String
    Dim bad As Variant, x As Variant, base As String, testName As String, n As Long
    s = Trim$(s)
    bad = Array("/", "\", "?", "*", "[", "]", ":")
    For Each x In bad: s = Replace(s, CStr(x), "-"): Next x
    s = Replace(s, "'", "")
    If Len(s) = 0 Then s = "Unassigned"
    If Len(s) > 31 Then s = Left$(s, 31)
    base = s: testName = base: n = 2
    Do While SheetExists(testName, wb)
        testName = Left$(base, 27) & " (" & CStr(n) & ")"
        n = n + 1
    Loop
    SafeSheetName = testName
End Function

Private Sub SortManagersByLastName(ByRef arr As Variant)
    Dim i As Long, j As Long, tmp As Variant
    On Error Resume Next
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(ManagerSortKey(CStr(arr(i))), ManagerSortKey(CStr(arr(j))), vbTextCompare) > 0 Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
    On Error GoTo 0
End Sub

Private Function ManagerSortKey(ByVal fullName As String) As String
    Dim p As Long
    fullName = Trim$(fullName)
    p = InStrRev(fullName, " ")
    If p > 0 Then ManagerSortKey = Mid$(fullName, p + 1) & "|" & Left$(fullName, p - 1) Else ManagerSortKey = fullName
End Function

Private Sub SortTextArray(ByRef arr As Variant)
    Dim i As Long, j As Long, tmp As Variant
    On Error Resume Next
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(CStr(arr(i)), CStr(arr(j)), vbTextCompare) > 0 Then tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
        Next j
    Next i
    On Error GoTo 0
End Sub

' ====================== WORKBOOK STATE ======================
' Ensures the master is saved as .xlsm so the VBA module is retained; finds the prior report's source sheet.
Private Function EnsureMacroEnabledWorkbook() As Boolean
    Dim savePath As Variant, baseName As String
    Dim isMacroEnabled As Boolean
    On Error GoTo Fail

    isMacroEnabled = (ThisWorkbook.FileFormat = xlOpenXMLWorkbookMacroEnabled) And _
                     (LCase$(Right$(ThisWorkbook.Name, 5)) = ".xlsm")
    If isMacroEnabled Then
        EnsureMacroEnabledWorkbook = True
        Exit Function
    End If

    baseName = ThisWorkbook.Name
    If InStrRev(baseName, ".") > 0 Then baseName = Left$(baseName, InStrRev(baseName, ".") - 1)

    MsgBox "This master workbook is currently an .xlsx workbook." & vbCrLf & vbCrLf & _
           "Because the Pending Pickup builder is VBA-based, the master must first be saved as an Excel Macro-Enabled Workbook (.xlsm)." & vbCrLf & vbCrLf & _
           "Choose the .xlsm filename in the Save As window that opens next.", vbInformation, "Pending Pickup - Save Master as .xlsm"

    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=baseName & ".xlsm", _
        FileFilter:="Excel Macro-Enabled Workbook (*.xlsm),*.xlsm", _
        Title:="Save Pending Pickup master as .xlsm")

    If VarType(savePath) = vbBoolean Then
        If savePath = False Then
            MsgBox "Save As was cancelled. The Pending Pickup report was not run.", vbExclamation
            Exit Function
        End If
    End If

    If LCase$(Right$(CStr(savePath), 5)) <> ".xlsm" Then savePath = CStr(savePath) & ".xlsm"

    Application.DisplayAlerts = False
    ThisWorkbook.SaveAs Filename:=CStr(savePath), FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    Application.DisplayAlerts = True

    isMacroEnabled = (ThisWorkbook.FileFormat = xlOpenXMLWorkbookMacroEnabled) And _
                     (LCase$(Right$(ThisWorkbook.Name, 5)) = ".xlsm")
    If Not isMacroEnabled Then
        MsgBox "Excel did not switch the master workbook to .xlsm. Please save the master manually as an Excel Macro-Enabled Workbook and run the report again.", vbCritical
        Exit Function
    End If

    EnsureMacroEnabledWorkbook = True
    Exit Function

Fail:
    Application.DisplayAlerts = True
    MsgBox "I could not save the master workbook as .xlsm." & vbCrLf & Err.Description, vbCritical
End Function

Private Function FindPriorSourceSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    Dim headerRow As Long

    If SheetExists(RAW_SHEET, wb) Then
        Set ws = wb.Worksheets(RAW_SHEET)
        headerRow = FindHeaderRow(ws)
        If headerRow > 0 Then
            Set FindPriorSourceSheet = ws
            Exit Function
        End If
    End If

    For Each ws In wb.Worksheets
        headerRow = FindHeaderRow(ws)
        If headerRow > 0 Then
            Set FindPriorSourceSheet = ws
            Exit Function
        End If
    Next ws
End Function

' ====================== DIALOG / FILE / FOLDER ======================
' File/folder pickers, remembered-folder persistence, prior-report discovery, and output path building.
Private Function SelectExcelFile(ByVal dialogTitle As String, ByVal initialFolder As String, ByVal suggestedFile As String, ByVal allowCancel As Boolean) As Variant
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
    Dim folder As String, sep As Long
    sep = InStrRev(filePath, Application.PathSeparator)
    If sep > 0 Then folder = Left$(filePath, sep - 1) Else folder = filePath
    On Error Resume Next
    SaveSetting "PendingPickupReport", "Settings", settingName, folder
    On Error GoTo 0
End Sub

Private Function GetSavedFolder(ByVal settingName As String) As String
    On Error Resume Next
    GetSavedFolder = GetSetting("PendingPickupReport", "Settings", settingName, "")
    On Error GoTo 0
End Function

Private Function FindMostRecentPendingPickupReport(ByVal folder As String, ByVal masterPath As String) As String
    Dim f As String, fullPath As String, bestPath As String
    Dim bestDate As Date, candidateDate As Date

    If Len(folder) = 0 Then Exit Function
    If Len(Dir(folder, vbDirectory)) = 0 Then Exit Function

    f = Dir(folder & Application.PathSeparator & "Pending Pickup *.xlsx")
    Do While Len(f) > 0
        fullPath = folder & Application.PathSeparator & f
        If StrComp(fullPath, masterPath, vbTextCompare) <> 0 Then
            candidateDate = FileDateTime(fullPath)
            If Len(bestPath) = 0 Or candidateDate > bestDate Then
                bestDate = candidateDate
                bestPath = fullPath
            End If
        End If
        f = Dir()
    Loop
    FindMostRecentPendingPickupReport = bestPath
End Function

Private Function BuildOutputPath(srcWb As Workbook) As String
    Dim folder As String, baseName As String, candidate As Variant
    On Error GoTo Fail
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
    On Error Resume Next
    GetSavedOutputFolder = GetSetting("PendingPickupReport", "Settings", "OutputFolder", "")
    On Error GoTo 0
End Function

Private Sub SaveOutputFolder(ByVal folder As String)
    On Error Resume Next
    SaveSetting "PendingPickupReport", "Settings", "OutputFolder", folder
    On Error GoTo 0
End Sub



' v3.4 final change-log note: Manager raw-data header moved to row 10 to prevent KPI overlap.
