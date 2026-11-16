Attribute VB_Name = "modCore"
Option Explicit

' ============================================================================
'  Module: modCore
'  Owns the full report build: source selection, prior comparison, manager tabs,
'  dashboard, Raw Data copy, and save. Public entry point wired to the button.
' ============================================================================

Public Function CreatePendingPickupReport() As Boolean
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
    Dim comparisonEnabled As Boolean
    Dim errText As String
    Dim managerLastCol As Long, managerPickupDateCol As Long, managerStartDateCol As Long, managerAgeCol As Long, managerRepCol As Long
    Dim sortedManagerRows As Collection
    Dim reportStamp As String, sourceName As String
    Dim managerSheetNames As Object
    Dim keyCol As Long
    Dim dtCheck As Date, repName As String, ck As String, pk As String, ck2 As String
    Dim rowItem As Variant

    Set managerSheetNames = CreateObject("Scripting.Dictionary")
    managerSheetNames.CompareMode = vbTextCompare

    CreatePendingPickupReport = False
    mInReportRun = True
    mReportStartTime = Now

    On Error GoTo CleanFail

    mRunLogPath = BuildRunLogPath
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    LogMsg "=== CreatePendingPickupReport started ==="
    LogMsg "Run log: " & mRunLogPath
    LogMsg "Completion marker: " & COMPLETION_MARKER_PATH
    LogMsg "Excel version: " & Application.Version
    LogMsg "Start time: " & Format(mReportStartTime, "yyyy-mm-dd hh:n:ss")

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
    If Not EnsureMacroEnabledWorkbook Then GoTo CleanExit

    ' Select the CURRENT source report. Nothing needs to be pasted into the master.
    Application.StatusBar = "Select the current All Open Treatment Plans report..."
    Dim currentPath As Variant

    ' Auto mode: read source path from file if present (for automated runs).
    Dim autoSourcePath As String
    autoSourcePath = ""
    On Error Resume Next
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(AUTO_SOURCE_PATH_FILE) Then
        Set ts = fso.OpenTextFile(AUTO_SOURCE_PATH_FILE, 1)
        autoSourcePath = ts.ReadAll
        ts.Close
        Set ts = Nothing
    End If
    Set fso = Nothing
    On Error GoTo CleanFail

    If Len(autoSourcePath) > 0 Then
        currentPath = autoSourcePath
        LogMsg "AUTO: Using source path from file: " & currentPath
    Else
        currentPath = SelectExcelFile( _
            "Select the CURRENT All Open Treatment Plans report", _
            GetSavedFolder("CurrentFolder"), _
            "", _
            False)
    End If

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
    keyCol = FindHeaderColumn(raw, headerRow, KEY_HEADER)

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
            repName = Trim$(CStr(currentData(r, repCol)))
            If Len(repName) > 0 Then
                If Not mRepSeenDict.Exists(repName) Then
                    mRepSeenDict.Add repName, r
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

    ' Auto mode: skip PRIOR picker when running unattended.
    Dim autoSkipPrior As Boolean
    autoSkipPrior = False
    On Error Resume Next
    Dim fsoPrior As Object
    Set fsoPrior = CreateObject("Scripting.FileSystemObject")
    If fsoPrior.FileExists(AUTO_SKIP_PRIOR_FILE) Then
        autoSkipPrior = True
    End If
    Set fsoPrior = Nothing
    On Error GoTo CleanFail

    If autoSkipPrior Then
        priorPath = False
        LogMsg "AUTO: Skipping PRIOR picker (unattended run)."
    Else
        priorPath = SelectExcelFile( _
            "Select the PRIOR Pending Pickup report (Cancel for no comparison)", _
            GetSavedFolder("PriorFolder"), _
            priorSuggested, _
            True)
    End If

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
            End If
        Next r
    End If

    priorCount = priorRows.Count
    stillCount = 0
    newCount = 0
    clearedCount = 0
    LogMsg "priorCount=" & priorCount & " hasPrior=" & hasPrior

    ' Compare current vs prior to classify new/cleared/still
    comparisonEnabled = (keyCol > 0 And nameCol > 0)
    If Not comparisonEnabled Then
        LogMsg "TRACE: Skipping prior comparison (keyCol=" & keyCol & " nameCol=" & nameCol & ")"
    End If
    Set currentByKey = CreateObject("Scripting.Dictionary")
    currentByKey.CompareMode = vbTextCompare
    Set priorByKey = CreateObject("Scripting.Dictionary")
    priorByKey.CompareMode = vbTextCompare
    LogMsg "TRACE: Starting comparison block keyCol=" & keyCol & " nameCol=" & nameCol & " currentRows=" & currentRows.Count & " priorRows=" & priorRows.Count

    If comparisonEnabled Then
        On Error Resume Next
        For Each rowItem In currentRows
            ck = GetRecordKey(currentData, CLng(rowItem), keyCol, nameCol)
            If Err.Number <> 0 Then LogMsg "TRACE: currentByKey Err=" & Err.Number & " row=" & rowItem
            Err.Clear
            If Len(ck) > 0 Then currentByKey(ck) = CLng(rowItem)
        Next rowItem

        For Each rowItem In priorRows
            pk = GetRecordKey(priorData, CLng(rowItem), priorStatusCol, priorNameCol)
            If Err.Number <> 0 Then LogMsg "TRACE: priorByKey Err=" & Err.Number & " row=" & rowItem
            Err.Clear
            If Len(pk) > 0 Then
                priorByKey(pk) = CLng(rowItem)
                If currentByKey.Exists(pk) Then
                    stillCount = stillCount + 1
                Else
                    clearedCount = clearedCount + 1
                End If
            End If
        Next rowItem

        For Each rowItem In currentRows
            ck2 = GetRecordKey(currentData, CLng(rowItem), keyCol, nameCol)
            If Err.Number <> 0 Then LogMsg "TRACE: newCount Err=" & Err.Number & " row=" & rowItem
            Err.Clear
            If Len(ck2) > 0 Then
                If Not priorByKey.Exists(ck2) Then
                    newCount = newCount + 1
                End If
            End If
        Next rowItem
        On Error GoTo CleanFail
    End If

    CalculateAgingStats currentData, currentRows, pickupDateCol, age30Count, age60Count, age90Count, missingDateCount
    LogMsg "TRACE: CalculateAgingStats done"
    LogMsg "Aging: 30=" & age30Count & " 60=" & age60Count & " 90=" & age90Count & " missing=" & missingDateCount
    AuditLogEvent "AGING", "30=" & age30Count & " 60=" & age60Count & " 90=" & age90Count & " missing=" & missingDateCount
    mAuditMissingDates = missingDateCount

    raw.Range(raw.Cells(headerRow, 1), raw.Cells(lastRow, lastCol)).Columns.AutoFit
    ApplySourceDateFormats raw, headerRow, lastRow, startDateCol, pickupDateCol
    LogMsg "TRACE: ApplySourceDateFormats done"

    Set outWb = Workbooks.Add(xlWBATWorksheet)
    LogMsg "TRACE: Workbooks.Add done"
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
        LogMsg "  worksheet added"
        ws.Name = SafeSheetName(manager, outWb)
        managerSheetNames(manager) = ws.Name
        LogMsg "  sheet named: " & ws.Name

        Set sortedManagerRows = SortManagerRowsBySalesRep(currentData, managerRows(manager), repCol, nameCol)
        LogMsg "  rows sorted, count=" & sortedManagerRows.Count
        arrOut = BuildManagerOutputArray(currentData, sortedManagerRows, lastCol, pickupDateCol)
        LogMsg "  output array built: " & UBound(arrOut, 1) & "x" & UBound(arrOut, 2)
        managerLastCol = UBound(arrOut, 2)
        ws.Range(ws.Cells(MANAGER_HEADER_ROW, 1), ws.Cells(MANAGER_HEADER_ROW + UBound(arrOut, 1) - 1, managerLastCol)).Value = arrOut
        LogMsg "  values written"
        managerPickupDateCol = FindHeaderColumn(ws, MANAGER_HEADER_ROW, PICKUP_DATE_HEADER)
        managerStartDateCol = FindHeaderColumn(ws, MANAGER_HEADER_ROW, START_DATE_HEADER)
        managerAgeCol = FindHeaderColumn(ws, MANAGER_HEADER_ROW, "Pickup Age (Days)")
        managerRepCol = FindHeaderColumn(ws, MANAGER_HEADER_ROW, REP_HEADER)
        FormatManagerSheet ws, manager, managerRows(manager).Count, managerLastCol, managerPickupDateCol, managerStartDateCol, managerAgeCol, managerRepCol, MANAGER_HEADER_ROW + 1, LastUsedRow(ws), managerAgeCol, managerPickupDateCol, currentData, sortedManagerRows, hasPrior, pickupDateCol
        ApplyDateFormatsToOutput ws, MANAGER_HEADER_ROW, LastUsedRow(ws), managerStartDateCol, managerPickupDateCol
        LogMsg "  manager tab done: " & manager

        ' Add email button to the right of manager data
        Dim emailBtn As Shape
        Set emailBtn = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
            ws.Columns(managerLastCol + 1).Left + 10, _
            ws.Rows(MANAGER_HEADER_ROW).Top, _
            130, 35)
        emailBtn.Name = "btnEmail" & SafeSheetName(manager, outWb)
        emailBtn.TextFrame2.TextRange.Text = "Email " & manager
        emailBtn.TextFrame2.TextRange.Font.Bold = True
        emailBtn.TextFrame2.TextRange.Font.Size = 10
        emailBtn.OnAction = "EmailSelectedManagerTab"
        emailBtn.AlternativeText = ws.Name
    Next i
    LogMsg "All manager tabs done"

    ' Apply uniform column widths across all manager tabs (based on the longest content in each column)
    ApplyUniformManagerColumnWidths outWb, managers, managerSheetNames

    ' Build the dashboard after manager sheets exist so manager names can link directly to their tabs.
    LogMsg "BuildDashboard start"
    BuildDashboard dashboard, currentCount, priorCount, newCount, clearedCount, stillCount, _
               age30Count, age60Count, age90Count, missingDateCount, hasPrior, managers, managerRows, managerSheetNames, currentData, keyCol, nameCol, pickupDateCol, _
               reportStamp, sourceName, currentRows
    LogMsg "BuildDashboard done"

    ' Add email dashboard button
    Dim dashBtn As Shape
    Set dashBtn = dashboard.Shapes.AddShape(msoShapeRoundedRectangle, _
        dashboard.Columns("Q").Left + 10, _
        dashboard.Rows(2).Top, _
        170, 35)
    dashBtn.Name = "btnEmailDashboard"
    dashBtn.TextFrame2.TextRange.Text = "Email Dashboard to Director"
    dashBtn.TextFrame2.TextRange.Font.Bold = True
    dashBtn.TextFrame2.TextRange.Font.Size = 10
    dashBtn.OnAction = "EmailDashboard"
    dashBtn.AlternativeText = dashboard.Name

    Set rawOut = outWb.Worksheets.Add(After:=outWb.Worksheets(outWb.Worksheets.Count))
    rawOut.Name = "Raw Data"
    LogMsg "Raw Data sheet created"
    LogMsg "TRACE: BuildOutputArray start"
    arrRaw = BuildOutputArray(currentData, currentRows, lastCol)
    LogMsg "TRACE: BuildOutputArray done, rows=" & UBound(arrRaw, 1) & " cols=" & UBound(arrRaw, 2)
    LogMsg "TRACE: writing raw data values to sheet"
    rawOut.Range(rawOut.Cells(1, 1), rawOut.Cells(UBound(arrRaw, 1), lastCol)).Value = arrRaw
    LogMsg "TRACE: raw data values written"
    LogMsg "TRACE: FormatRawData start"
    FormatRawData rawOut, lastCol, UBound(arrRaw, 1)
    LogMsg "TRACE: FormatRawData done"
    LogMsg "TRACE: CopyColumnFormats start"
    CopyColumnFormats raw, rawOut, lastCol
    LogMsg "TRACE: CopyColumnFormats done"
    LogMsg "TRACE: MatchSourceColumnWidths start"
    MatchSourceColumnWidths raw, rawOut, lastCol
    LogMsg "TRACE: MatchSourceColumnWidths done"
    LogMsg "TRACE: ApplyDateFormatsToOutput start"
    ApplyDateFormatsToOutput rawOut, 1, UBound(arrRaw, 1), startDateCol, pickupDateCol
    LogMsg "TRACE: ApplyDateFormatsToOutput done"
    LogMsg "TRACE: ApplyDateFormatByHeader (Patient Date of Birth) start"
    ApplyDateFormatByHeader rawOut, 1, UBound(arrRaw, 1), "Patient Date of Birth"
    LogMsg "TRACE: ApplyDateFormatByHeader (Patient Date of Birth) done"
    LogMsg "TRACE: ApplyDateFormatByHeader (Last Modified) start"
    ApplyDateFormatByHeader rawOut, 1, UBound(arrRaw, 1), "Last Modified"
    LogMsg "TRACE: ApplyDateFormatByHeader (Last Modified) done"

    Application.StatusBar = "Saving finished Pending Pickup report..."
    savePath = BuildOutputPath(srcWb)
    LogMsg "TRACE: BuildOutputPath done, savePath=" & savePath
    LogMsg "Saving to: " & savePath

    ' Hide unused rows/columns on every sheet so the report looks clean
    Dim hideWs As Worksheet
    Dim hideCount As Long
    hideCount = 0
    For Each hideWs In outWb.Worksheets
        HideUnusedCells hideWs
        hideCount = hideCount + 1
    Next hideWs
    LogMsg "TRACE: HideUnusedCells done for " & hideCount & " sheets"

    ' Write audit data to the output workbook
    mAuditOutputPath = savePath
    LogMsg "TRACE: AuditWriteSheet start"
    AuditWriteSheet outWb
    LogMsg "TRACE: AuditWriteSheet done"
    LogMsg "TRACE: AuditLogEvent SAVE start"
    AuditLogEvent "SAVE", "Report saved to: " & savePath
    LogMsg "TRACE: AuditLogEvent SAVE done"

    LogMsg "TRACE: SaveAs start"
    outWb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    LogMsg "TRACE: SaveAs done"
    LogMsg "Save done"
    Dim runDuration As Double
    runDuration = TimerDiff(mReportStartTime, Now)
    LogMsg "Run duration (seconds): " & Format(runDuration, "0.00")
    WriteCompletionMarker savePath
    LogMsg "Completion marker written: " & COMPLETION_MARKER_PATH
    LogMsg "Output file: " & savePath
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
    LogErr "CreatePendingPickupReport"
    Application.StatusBar = False
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.Calculation = oldCalc
    mInReportRun = False
    mRunLogPath = ""
    If Not outWb Is Nothing Then outWb.Close SaveChanges:=False
    If Not priorWb Is Nothing Then priorWb.Close SaveChanges:=False
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
    On Error GoTo 0
    LogMsg "=== CleanFail exit ==="
    CreatePendingPickupReport = False
    WriteCompletionMarker "(failed)"
    LogMsg "Completion marker written (failure): " & COMPLETION_MARKER_PATH
End Function
