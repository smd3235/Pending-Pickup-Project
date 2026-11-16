Attribute VB_Name = "modUtils"
Option Explicit

' ============================================================================
'  Module: modUtils
'  General utility functions: sheet/header detection, date handling,
'  record keys, aging stats, output arrays, sorting, formatting helpers,
'  column width matching, hiding unused cells.
' ============================================================================

' ----------------------------------------------------------------------------
' Sheet detection
' ----------------------------------------------------------------------------

Public Function FindPriorSourceSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim foundCols As Long

    On Error Resume Next
    For Each ws In wb.Worksheets
        headerRow = FindHeaderRow(ws)
        If headerRow > 0 Then
            foundCols = 0
            If FindHeaderColumn(ws, headerRow, MANAGER_HEADER) > 0 Then foundCols = foundCols + 1
            If FindHeaderColumn(ws, headerRow, STATUS_HEADER) > 0 Then foundCols = foundCols + 1
            If FindHeaderColumn(ws, headerRow, REP_HEADER) > 0 Then foundCols = foundCols + 1
            If FindHeaderColumn(ws, headerRow, NAME_HEADER) > 0 Then foundCols = foundCols + 1
            If foundCols >= 3 Then
                Set FindPriorSourceSheet = ws
                Exit Function
            End If
        End If
    Next ws
    Set FindPriorSourceSheet = Nothing
End Function

Public Function FindHeaderRow(ByVal ws As Worksheet) As Long
    Dim lastRow As Long, r As Long
    Dim cnt As Long, hdr As String

    On Error Resume Next
    lastRow = LastUsedRow(ws)
    FindHeaderRow = 0

    For r = 1 To lastRow
        cnt = 0
        hdr = LCase$(Trim$(CStr(ws.Cells(r, 1).Value)))
        If hdr = LCase$(MANAGER_HEADER) Or hdr = LCase$(STATUS_HEADER) Or hdr = LCase$(REP_HEADER) Or hdr = LCase$(NAME_HEADER) Then cnt = cnt + 1
        If FindHeaderColumn(ws, r, MANAGER_HEADER) > 0 Then cnt = cnt + 1
        If FindHeaderColumn(ws, r, STATUS_HEADER) > 0 Then cnt = cnt + 1
        If FindHeaderColumn(ws, r, REP_HEADER) > 0 Then cnt = cnt + 1
        If FindHeaderColumn(ws, r, NAME_HEADER) > 0 Then cnt = cnt + 1
        If cnt >= 3 Then
            FindHeaderRow = r
            Exit Function
        End If
    Next r
End Function

Public Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerText As String) As Long
    Dim c As Long, lastCol As Long, h As String

    On Error Resume Next
    FindHeaderColumn = 0
    lastCol = LastUsedColumn(ws)

    For c = 1 To lastCol
        h = LCase$(Trim$(CStr(ws.Cells(headerRow, c).Value)))
        If h = LCase$(headerText) Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
End Function

Public Function LastUsedRow(ByVal ws As Worksheet) As Long
    On Error Resume Next
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        LastUsedRow = 0
        Exit Function
    End If
    LastUsedRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
End Function

Public Function LastUsedColumn(ByVal ws As Worksheet) As Long
    On Error Resume Next
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        LastUsedColumn = 0
        Exit Function
    End If
    LastUsedColumn = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
End Function

' ----------------------------------------------------------------------------
' Date handling
' ----------------------------------------------------------------------------

Public Function TryGetExcelDate(ByVal v As Variant, ByRef dt As Date) As Boolean
    On Error Resume Next
    TryGetExcelDate = False

    If IsEmpty(v) Or Len(Trim$(CStr(v))) = 0 Then Exit Function

    If IsDate(v) Then
        dt = CDate(v)
        TryGetExcelDate = True
        Exit Function
    End If

    If IsNumeric(v) Then
        dt = CDbl(v)
        TryGetExcelDate = True
        Exit Function
    End If

    dt = CDate(CStr(v))
    TryGetExcelDate = True
End Function

Public Sub NormalizeDateColumn(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long, ByVal col As Long)
    Dim r As Long, v As Variant, dt As Date, s As String

    On Error Resume Next
    If col <= 0 Or headerRow <= 0 Or lastRow <= headerRow Then Exit Sub

    For r = headerRow + 1 To lastRow
        v = ws.Cells(r, col).Value
        If IsEmpty(v) Then GoTo NextRow
        s = Trim$(CStr(v))
        If Len(s) = 0 Then GoTo NextRow

        If TryGetExcelDate(v, dt) Then
            ws.Cells(r, col).Value = dt
            ws.Cells(r, col).NumberFormat = "m/d/yyyy"
        Else
            ws.Cells(r, col).Value = s
        End If
NextRow:
    Next r
End Sub

Public Sub ApplySourceDateFormats(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long, ByVal startDateCol As Long, ByVal pickupDateCol As Long)
    Dim r As Long
    On Error Resume Next
    If startDateCol > 0 Then
        For r = headerRow + 1 To lastRow
            If IsDate(ws.Cells(r, startDateCol).Value) Then ws.Cells(r, startDateCol).NumberFormat = "m/d/yyyy"
        Next r
    End If
    If pickupDateCol > 0 Then
        For r = headerRow + 1 To lastRow
            If IsDate(ws.Cells(r, pickupDateCol).Value) Then ws.Cells(r, pickupDateCol).NumberFormat = "m/d/yyyy"
        Next r
    End If
End Sub

' ----------------------------------------------------------------------------
' Record keys
' ----------------------------------------------------------------------------

Public Function GetRecordKey(ByVal data As Variant, ByVal rowIdx As Long, ByVal keyCol As Long, ByVal nameCol As Long) As String
    Dim k As String, n As String
    On Error Resume Next
    GetRecordKey = ""
    If keyCol <= 0 Or nameCol <= 0 Then Exit Function
    k = Trim$(CStr(data(rowIdx, keyCol)))
    n = Trim$(CStr(data(rowIdx, nameCol)))
    GetRecordKey = "|" & k & "|" & n & "|"
End Function

' ----------------------------------------------------------------------------
' Aging stats
' ----------------------------------------------------------------------------

Public Sub CalculateAgingStats(ByVal data As Variant, ByVal rows As Collection, ByVal pickupDateCol As Long, _
    ByRef age30Count As Long, ByRef age60Count As Long, ByRef age90Count As Long, ByRef missingDateCount As Long)

    Dim r As Long, v As Variant, dt As Date
    Dim age As Long
    Dim item As Variant

    On Error Resume Next
    age30Count = 0
    age60Count = 0
    age90Count = 0
    missingDateCount = 0

    For Each item In rows
        r = CLng(item)
        v = data(r, pickupDateCol)
        If Len(Trim$(CStr(v))) = 0 Or Not TryGetExcelDate(v, dt) Then
            missingDateCount = missingDateCount + 1
        Else
            age = DateDiff("d", dt, Date)
            If age >= 90 Then
                age90Count = age90Count + 1
            ElseIf age >= 60 Then
                age60Count = age60Count + 1
            ElseIf age >= 30 Then
                age30Count = age30Count + 1
            End If
        End If
    Next item
End Sub

Public Function CalculateAverageAge(ByVal data As Variant, ByVal rows As Collection, ByVal pickupDateCol As Long) As Variant
    Dim r As Long, v As Variant, dt As Date
    Dim totalDays As Double, cnt As Long
    Dim item As Variant

    On Error Resume Next
    totalDays = 0
    cnt = 0

    For Each item In rows
        r = CLng(item)
        v = data(r, pickupDateCol)
        If Len(Trim$(CStr(v))) > 0 And TryGetExcelDate(v, dt) Then
            totalDays = totalDays + DateDiff("d", dt, Date)
            cnt = cnt + 1
        End If
    Next item

    If cnt > 0 Then
        CalculateAverageAge = totalDays / cnt
    Else
        CalculateAverageAge = "N/A"
    End If
End Function

' ----------------------------------------------------------------------------
' Output arrays
' ----------------------------------------------------------------------------

Public Function BuildOutputArray(ByVal data As Variant, ByVal rows As Collection, ByVal lastCol As Long) As Variant
    Dim arr() As Variant
    Dim r As Long, c As Long, i As Long
    Dim item As Variant

    ReDim arr(1 To rows.Count + 1, 1 To lastCol)

    ' Header row
    For c = 1 To lastCol
        arr(1, c) = data(1, c)
    Next c

    i = 2
    For Each item In rows
        r = CLng(item)
        For c = 1 To lastCol
            arr(i, c) = data(r, c)
        Next c
        i = i + 1
    Next item

    BuildOutputArray = arr
End Function

Private Function IsExcludedManagerHeader(ByVal headerText As String) As Boolean
    Dim h As String
    h = LCase$(Trim$(headerText))
    IsExcludedManagerHeader = False
    Select Case h
        Case LCase$(REP_HEADER), LCase$(START_DATE_HEADER), LCase$(NAME_HEADER), _
             LCase$("Billing Customer"), LCase$("Prescribing Physician"), _
             LCase$("No Rental Equipment Provided"), LCase$(PICKUP_DATE_HEADER), _
             LCase$("Pickup Age (Days)")
            IsExcludedManagerHeader = False
        Case LCase$(MANAGER_HEADER), LCase$(STATUS_HEADER), LCase$(KEY_HEADER)
            IsExcludedManagerHeader = True
        Case Else
            IsExcludedManagerHeader = True
    End Select
End Function

Private Function ManagerColumnRank(ByVal headerText As String, ByVal sourceCol As Long) As Long
    Select Case LCase$(Trim$(headerText))
        Case LCase$(REP_HEADER): ManagerColumnRank = 1
        Case LCase$(START_DATE_HEADER): ManagerColumnRank = 2
        Case LCase$(NAME_HEADER): ManagerColumnRank = 3
        Case LCase$("Billing Customer"): ManagerColumnRank = 4
        Case LCase$("Prescribing Physician"): ManagerColumnRank = 5
        Case LCase$("No Rental Equipment Provided"): ManagerColumnRank = 6
        Case LCase$(PICKUP_DATE_HEADER): ManagerColumnRank = 7
        Case Else: ManagerColumnRank = 1000 + sourceCol
    End Select
End Function

Public Function BuildManagerOutputArray(ByVal data As Variant, ByVal rows As Collection, ByVal lastCol As Long, ByVal pickupDateCol As Long) As Variant
    Dim arr() As Variant
    Dim i As Long, r As Long, c As Long, outCol As Long
    Dim item As Variant
    Dim keepCols() As Long
    Dim keepCount As Long
    Dim ageColIdx As Long
    Dim j As Long, tmp As Long, keyI As Long, keyJ As Long

    ' Select which source columns to keep and in what order
    keepCount = 0
    For c = 1 To lastCol
        If Not IsExcludedManagerHeader(CStr(data(1, c))) Then
            keepCount = keepCount + 1
            ReDim Preserve keepCols(1 To keepCount)
            keepCols(keepCount) = c
        End If
    Next c

    ' Sort kept columns by rank
    If keepCount > 1 Then
        For i = 1 To keepCount - 1
            For j = i + 1 To keepCount
                keyI = ManagerColumnRank(CStr(data(1, keepCols(i))), keepCols(i))
                keyJ = ManagerColumnRank(CStr(data(1, keepCols(j))), keepCols(j))
                If keyI > keyJ Then
                    tmp = keepCols(i): keepCols(i) = keepCols(j): keepCols(j) = tmp
                End If
            Next j
        Next i
    End If

    ' +1 for "Pickup Age (Days)" column
    ReDim arr(1 To rows.Count + 1, 1 To keepCount + 1)

    ' Header row
    For outCol = 1 To keepCount
        arr(1, outCol) = data(1, keepCols(outCol))
    Next outCol
    arr(1, keepCount + 1) = "Pickup Age (Days)"
    ageColIdx = keepCount + 1

    ' Data rows
    i = 2
    For Each item In rows
        r = CLng(item)
        For outCol = 1 To keepCount
            arr(i, outCol) = data(r, keepCols(outCol))
        Next outCol
        arr(i, ageColIdx) = "" ' formula written later by FormatManagerSheet context
        i = i + 1
    Next item

    BuildManagerOutputArray = arr
End Function

' ----------------------------------------------------------------------------
' Sorting
' ----------------------------------------------------------------------------

Public Sub SortManagersByLastName(ByRef names As Variant)
    Dim i As Long, j As Long
    Dim a As String, b As String
    Dim ta As String, tb As String
    Dim pa As Long, pb As Long

    On Error Resume Next

    For i = LBound(names) To UBound(names) - 1
        For j = i + 1 To UBound(names)
            a = CStr(names(i))
            b = CStr(names(j))
            pa = InStrRev(a, " ")
            pb = InStrRev(b, " ")
            ta = IIf(pa > 0, Trim$(Mid$(a, pa + 1)), LCase$(a))
            tb = IIf(pb > 0, Trim$(Mid$(b, pb + 1)), LCase$(b))
            If StrComp(ta, tb, vbTextCompare) > 0 Then
                names(i) = b
                names(j) = a
            End If
        Next j
    Next i
End Sub

Public Function SortManagerRowsBySalesRep(ByVal data As Variant, ByVal rows As Collection, ByVal repCol As Long, ByVal nameCol As Long) As Collection
    Dim sorted As New Collection
    Dim item As Variant, r As Long
    Dim rep As String
    Dim keyDict As Object

    Set keyDict = CreateObject("Scripting.Dictionary")
    keyDict.CompareMode = vbTextCompare

    For Each item In rows
        r = CLng(item)
        rep = Trim$(CStr(data(r, repCol)))
        If Len(rep) = 0 Then rep = "No Rep"
        keyDict.Add CStr(r), rep
    Next item

    ' Sort rows by rep name
    Dim keys As Variant
    keys = keyDict.Keys
    Dim i As Long, j As Long, k As String
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If StrComp(keyDict(keys(i)), keyDict(keys(j)), vbTextCompare) > 0 Then
                k = keys(i)
                keys(i) = keys(j)
                keys(j) = k
            End If
        Next j
    Next i

    For i = LBound(keys) To UBound(keys)
        sorted.Add CLng(keys(i))
    Next i

    Set SortManagerRowsBySalesRep = sorted
End Function

' ----------------------------------------------------------------------------
' Sheet name safety
' ----------------------------------------------------------------------------

Public Function SafeSheetName(ByVal name As String, ByVal wb As Workbook) As String
    Dim clean As String, c As String, i As Long
    Dim existing As Worksheet

    clean = name
    For i = 1 To Len(clean)
        c = Mid$(clean, i, 1)
        Select Case c
            Case "\", "/", "?", "*", "[", "]", ":"
                clean = Left$(clean, i - 1) & "_" & Mid$(clean, i + 1)
        End Select
    Next i
    If Len(clean) > 31 Then clean = Left$(clean, 31)

    ' Ensure uniqueness
    If clean = "" Then clean = "Manager"
    Dim base As String
    base = clean
    Dim n As Long
    n = 1
    On Error Resume Next
    Do
        Err.Clear
        Set existing = wb.Worksheets(clean)
        If Err.Number <> 0 Then
            Exit Do
        End If
        clean = base & "_" & n
        n = n + 1
    Loop

    SafeSheetName = clean
End Function

' ----------------------------------------------------------------------------
' Date format helpers
' ----------------------------------------------------------------------------

Public Sub ApplyDateFormatsToOutput(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal startDateCol As Long, ByVal pickupDateCol As Long)
    Dim r As Long

    On Error Resume Next
    If startDateCol > 0 Then
        For r = firstRow To lastRow
            If IsDate(ws.Cells(r, startDateCol).Value) Then ws.Cells(r, startDateCol).NumberFormat = "m/d/yyyy"
        Next r
    End If
    If pickupDateCol > 0 Then
        For r = firstRow To lastRow
            If IsDate(ws.Cells(r, pickupDateCol).Value) Then ws.Cells(r, pickupDateCol).NumberFormat = "m/d/yyyy"
        Next r
    End If
End Sub

Public Sub ApplyDateFormatByHeader(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long, ByVal headerText As String)
    Dim c As Long, r As Long, h As String
    Dim col As Long

    On Error Resume Next
    col = FindHeaderColumn(ws, headerRow, headerText)
    If col = 0 Then Exit Sub

    For r = headerRow + 1 To lastRow
        If IsDate(ws.Cells(r, col).Value) Then ws.Cells(r, col).NumberFormat = "m/d/yyyy"
    Next r
End Sub

' ----------------------------------------------------------------------------
' Column width matching
' ----------------------------------------------------------------------------

Public Sub MatchSourceColumnWidths(ByVal src As Worksheet, ByVal dst As Worksheet, ByVal lastCol As Long)
    Dim c As Long, w As Double

    On Error Resume Next
    For c = 1 To lastCol
        w = src.Columns(c).ColumnWidth
        If w > 0 Then dst.Columns(c).ColumnWidth = w
    Next c
End Sub

Public Sub MatchSourceColumnWidthsForManager(ByVal src As Worksheet, ByVal dst As Worksheet, ByVal lastCol As Long)
    MatchSourceColumnWidths src, dst, lastCol
End Sub

Public Sub CopyColumnFormats(ByVal src As Worksheet, ByVal dst As Worksheet, ByVal lastCol As Long)
    Dim c As Long

    On Error Resume Next
    ' Set column widths from source (fast, no clipboard)
    For c = 1 To lastCol
        dst.Columns(c).ColumnWidth = src.Columns(c).ColumnWidth
    Next c

    ' Copy header row formats in a single clipboard operation
    src.Range(src.Cells(1, 1), src.Cells(1, lastCol)).Copy
    dst.Range(dst.Cells(1, 1), dst.Cells(1, lastCol)).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
End Sub

' ----------------------------------------------------------------------------
' Hide unused cells
' ----------------------------------------------------------------------------

Public Sub HideUnusedCells(ByVal ws As Worksheet)
    Dim lastRow As Long, lastCol As Long
    Dim maxRow As Long, maxCol As Long
    Dim hideFromRow As Long, hideFromCol As Long

    On Error Resume Next

    lastRow = LastUsedRow(ws)
    lastCol = LastUsedColumn(ws)

    ' Hide unused columns to the right of data
    hideFromCol = lastCol + 1
    If hideFromCol <= ws.Columns.Count Then
        ws.Columns(hideFromCol & ":" & ws.Columns.Count).Hidden = True
    End If

    ' Hide unused rows below data
    hideFromRow = lastRow + 1
    If hideFromRow <= ws.Rows.Count Then
        ws.Rows(hideFromRow & ":" & ws.Rows.Count).Hidden = True
    End If
End Sub

' ----------------------------------------------------------------------------
' Timing helpers
' ----------------------------------------------------------------------------

Public Function TimerDiff(ByVal startTime As Date, ByVal endTime As Date) As Double
    TimerDiff = (endTime - startTime) * 86400#
End Function
