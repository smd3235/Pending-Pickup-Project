Option Explicit

' ============================================================================
'  Module: modAudit
'  Centralised audit logging for each report run.
'  Records metadata, record counts, data quality issues, and manager breakdowns
'  both to the run log and to an Audit Log sheet in the output workbook.
' ============================================================================

Public Sub AuditInit(ByVal srcPath As String)
    On Error Resume Next
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
    mAuditErrorCount = 0
    Set mAuditIssues = New Collection
End Sub

Public Sub AuditLogEvent(ByVal eventType As String, ByVal msg As String)
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    If Err.Number <> 0 Then Exit Sub
    Open GetLogPath() For Append As #f
    If Err.Number <> 0 Then Exit Sub
    Print #f, Format(Now, "yyyy-mm-dd hh:n:ss") & " | AUDIT | " & eventType & " | " & msg
    Close #f

    If mInReportRun And Len(mRunLogPath) > 0 Then
        f = FreeFile
        If Err.Number <> 0 Then Exit Sub
        Open mRunLogPath For Append As #f
        If Err.Number <> 0 Then Exit Sub
        Print #f, Format(Now, "yyyy-mm-dd hh:n:ss") & " | AUDIT | " & eventType & " | " & msg
        Close #f
    End If

    If Len(msg) > 0 Then
        mAuditIssues.Add eventType & ": " & msg
    End If
End Sub

Public Sub AuditIncrementCounter(ByVal counterName As String)
    On Error Resume Next
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

Public Sub AuditWriteSheet(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim r As Long
    Dim item As Variant

    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets(AUDIT_SHEET_NAME).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.name = AUDIT_SHEET_NAME

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

    ' Audit details section
    r = 25
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

    For Each item In mAuditIssues
        ws.Cells(r, 1).Value = Split(item, ":")(0)
        ws.Cells(r, 2).Value = Mid(item, InStr(item, ":") + 2)
        r = r + 1
    Next item

    ws.Columns(1).AutoFit
    ws.Columns(2).AutoFit

    LogMsg "Audit sheet written"
End Sub