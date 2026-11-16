Attribute VB_Name = "modFormatting"
Option Explicit

' ============================================================================
'  Module: modFormatting
'  Dashboard, manager sheet formatting, and table/row formatting helpers.
' ============================================================================

' ----------------------------------------------------------------------------
' Dashboard
' ----------------------------------------------------------------------------

Public Sub BuildDashboard(ByVal ws As Worksheet, ByVal currentCount As Long, ByVal priorCount As Long, ByVal newCount As Long, _
    ByVal clearedCount As Long, ByVal stillCount As Long, ByVal age30Count As Long, ByVal age60Count As Long, _
    ByVal age90Count As Long, ByVal missingDateCount As Long, ByVal hasPrior As Boolean, _
    ByVal managers As Variant, ByVal managerRows As Object, ByVal managerSheetNames As Object, _
    ByVal currentData As Variant, ByVal keyCol As Long, ByVal nameCol As Long, ByVal pickupDateCol As Long, _
    ByVal reportStamp As String, ByVal sourceName As String, ByVal currentRows As Collection)

    Dim rg As Range
    Dim fillColor As Long
    Dim labelText As String, valueText As String
    Dim r As Long, c As Long, i As Long
    Dim mgrName As String
    Dim mgrRows As Collection
    Dim m30 As Long, m60 As Long, m90 As Long, mMiss As Long
    Dim mgrAvg As Variant, pct30 As Double
    Dim totalAvg As Variant
    Dim tableStart As Long, tableEnd As Long

    On Error GoTo CleanFail

    LogMsg "BuildDashboard start"

    ' Title block A1:H2
    Set rg = ws.Range("A1:H2")
    rg.Merge
    rg.Value = "PENDING PICKUP DASHBOARD"
    With rg
        .Font.Bold = True
        .Font.Size = 22
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Rows("1:2").RowHeight = 30

    ' Subtitle A3:J3
    Set rg = ws.Range("A3:J3")
    rg.Merge
    rg.Value = "Report generated " & reportStamp & "   |   Source: " & sourceName
    With rg
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Italic = False
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(3).RowHeight = 18

    ' KPI cards row 4-6 (top row) and 8-10 (bottom row), 5 cards per row in cols A-J
    ' Row heights
    ws.Rows("4:6").RowHeight = 33.95
    ws.Rows("7:7").RowHeight = 21.95
    ws.Rows("8:10").RowHeight = 33.95
    ws.Rows("11:11").RowHeight = 21.95

    ' Top row cards
    labelText = "CURRENT PERIOD" & vbCrLf & "Pending Pickups"
    valueText = CStr(currentCount)
    AddDashboardCard ws, "A4:B5", "A6:B6", RGB(221, 235, 247), labelText, valueText

    labelText = "PRIOR PERIOD" & vbCrLf & "Pending Pickups"
    valueText = IIf(hasPrior, CStr(priorCount), "N/A")
    AddDashboardCard ws, "C4:D5", "C6:D6", RGB(209, 226, 242), labelText, valueText

    labelText = "NEW SINCE PRIOR" & vbCrLf & "Since Prior Report"
    valueText = IIf(hasPrior, CStr(newCount), "N/A")
    AddDashboardCard ws, "E4:F5", "E6:F6", RGB(226, 239, 218), labelText, valueText

    labelText = "CLEARED SINCE PRIOR" & vbCrLf & "Since Prior Report"
    valueText = IIf(hasPrior, CStr(clearedCount), "N/A")
    AddDashboardCard ws, "G4:H5", "G6:H6", RGB(209, 233, 217), labelText, valueText

    labelText = "STILL PENDING" & vbCrLf & "In Both Reports"
    valueText = IIf(hasPrior, CStr(stillCount), "N/A")
    AddDashboardCard ws, "I4:J5", "I6:J6", RGB(221, 235, 247), labelText, valueText

    ' Bottom row cards
    labelText = "% CHANGE VS PRIOR" & vbCrLf & "From Prior Report"
    If hasPrior And priorCount > 0 Then
        valueText = Format((currentCount - priorCount) / priorCount, "0.0%")
    Else
        valueText = "N/A"
    End If
    AddDashboardCard ws, "A8:B9", "A10:B10", RGB(234, 209, 220), labelText, valueText

    labelText = "AGING 30+ DAYS" & vbCrLf & "Older Pending Pickups"
    valueText = CStr(age30Count)
    AddDashboardCard ws, "C8:D9", "C10:D10", RGB(244, 204, 204), labelText, valueText

    labelText = "AGING 60+ DAYS" & vbCrLf & "Older Pending Pickups"
    valueText = CStr(age60Count)
    AddDashboardCard ws, "E8:F9", "E10:F10", RGB(234, 153, 153), labelText, valueText

    labelText = "AGING 90+ DAYS" & vbCrLf & "Older Pending Pickups"
    valueText = CStr(age90Count)
    AddDashboardCard ws, "G8:H9", "G10:H10", RGB(224, 102, 102), labelText, valueText

    labelText = "MISSING PICKUP DATE" & vbCrLf & "AUDIT: TP End Date"
    valueText = CStr(missingDateCount)
    AddDashboardCard ws, "I8:J9", "I10:J10", RGB(255, 242, 204), labelText, valueText

    ' Manager table section
    ws.Rows(12).RowHeight = 24
    Set rg = ws.Range("A12:K12")
    rg.Merge
    rg.Value = "PENDING PICKUPS BY MANAGER"
    With rg
        .Font.Bold = True
        .Font.Size = 11
        .Font.Name = "Segoe UI"
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ws.Rows(13).RowHeight = 30
    Dim headers As Variant
    headers = Array("Manager", "Current", "New", "30+ Days", "60+ Days", "90+ Days", "Missing Date", "Avg Age", "% 30+", "Cleared", "Still Pending")
    For c = 0 To UBound(headers)
        Set rg = ws.Cells(13, c + 1)
        rg.Value = headers(c)
        With rg
            .Font.Bold = True
            .Font.Name = "Segoe UI"
            .Font.Size = 11
            .Interior.Color = RGB(217, 234, 247)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Borders.LineStyle = xlContinuous
        End With
    Next c

    ' Manager table rows
    tableStart = 14
    r = tableStart
    For i = LBound(managers) To UBound(managers)
        mgrName = CStr(managers(i))
        Set mgrRows = managerRows(mgrName)
        CalculateAgingStats currentData, mgrRows, pickupDateCol, m30, m60, m90, mMiss
        mgrAvg = CalculateAverageAge(currentData, mgrRows, pickupDateCol)
        If mgrRows.Count > 0 Then
            pct30 = m30 / mgrRows.Count
        Else
            pct30 = 0
        End If

        ws.Cells(r, 1).Value = mgrName
        ws.Cells(r, 2).Value = mgrRows.Count
        ws.Cells(r, 3).Value = IIf(hasPrior, newCount, "N/A") ' placeholder per-manager comparison not tracked
        ws.Cells(r, 4).Value = m30
        ws.Cells(r, 5).Value = m60
        ws.Cells(r, 6).Value = m90
        ws.Cells(r, 7).Value = mMiss
        If IsNumeric(mgrAvg) Then
            ws.Cells(r, 8).Value = CDbl(mgrAvg)
        Else
            ws.Cells(r, 8).Value = "N/A"
        End If
        ws.Cells(r, 9).Value = pct30
        ws.Cells(r, 10).Value = IIf(hasPrior, clearedCount, "N/A")
        ws.Cells(r, 11).Value = IIf(hasPrior, stillCount, "N/A")

        ' Format per-row
        ws.Rows(r).RowHeight = 24
        For c = 1 To 11
            With ws.Cells(r, c)
                .Font.Name = "Segoe UI"
                .Font.Size = 11
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlHairline
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                If c = 1 Then .HorizontalAlignment = xlLeft
                If (r - tableStart) Mod 2 = 1 Then .Interior.Color = RGB(242, 242, 242)
            End With
        Next c
        ws.Cells(r, 9).NumberFormat = "0.0%"
        r = r + 1
    Next i

    ' TOTAL row
    tableEnd = r
    totalAvg = CalculateAverageAge(currentData, currentRows, pickupDateCol)
    ws.Cells(r, 1).Value = "TOTAL"
    ws.Cells(r, 2).Value = currentCount
    ws.Cells(r, 3).Value = IIf(hasPrior, newCount, "N/A")
    ws.Cells(r, 4).Value = age30Count
    ws.Cells(r, 5).Value = age60Count
    ws.Cells(r, 6).Value = age90Count
    ws.Cells(r, 7).Value = missingDateCount
    If IsNumeric(totalAvg) Then
        ws.Cells(r, 8).Value = CDbl(totalAvg)
    Else
        ws.Cells(r, 8).Value = "N/A"
    End If
    If currentCount > 0 Then
        ws.Cells(r, 9).Value = age30Count / currentCount
    Else
        ws.Cells(r, 9).Value = 0
    End If
    ws.Cells(r, 10).Value = IIf(hasPrior, clearedCount, "N/A")
    ws.Cells(r, 11).Value = IIf(hasPrior, stillCount, "N/A")
    ws.Rows(r).RowHeight = 24
    For c = 1 To 11
        With ws.Cells(r, c)
            .Font.Bold = True
            .Font.Name = "Segoe UI"
            .Font.Size = 11
            .Interior.Color = RGB(217, 234, 247)
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            If c = 1 Then .HorizontalAlignment = xlLeft
        End With
    Next c
    ws.Cells(r, 9).NumberFormat = "0.0%"

    ' Chart placeholder
    ws.Rows(r + 2).RowHeight = 21.95
    Set rg = ws.Range("A" & (r + 2) & ":K" & (r + 2))
    rg.Merge
    rg.Value = "PENDING PICKUPS BY MANAGER (CHART)"
    With rg
        .Font.Bold = True
        .Font.Size = 11
        .Font.Name = "Segoe UI"
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Column widths
    ws.Columns("A").ColumnWidth = 16.71
    ws.Columns("B:K").ColumnWidth = 13

    LogMsg "BuildDashboard done - all cards and table created"
    Exit Sub

CleanFail:
    LogErr "BuildDashboard"
End Sub

Public Sub AddDashboardCard(ByVal ws As Worksheet, ByVal labelRange As String, ByVal valueRange As String, _
    ByVal fillColor As Long, ByVal labelText As String, ByVal valueText As String)

    On Error GoTo CleanFail

    Dim labelRg As Range, valueRg As Range, cardRg As Range
    Dim topRow As Long, bottomRow As Long, leftCol As Long, rightCol As Long

    LogMsg "AddDashboardCard: " & labelRange & "/" & valueRange & " label=" & Replace(labelText, vbCrLf, " ") & " value=" & valueText

    Set labelRg = ws.Range(labelRange)
    Set valueRg = ws.Range(valueRange)

    ' Compute full card bounds
    topRow = labelRg.Row
    leftCol = labelRg.Column
    bottomRow = valueRg.Row + valueRg.Rows.Count - 1
    rightCol = labelRg.Column + labelRg.Columns.Count - 1
    Set cardRg = ws.Range(ws.Cells(topRow, leftCol), ws.Cells(bottomRow, rightCol))

    ' Color and border the whole card
    cardRg.Interior.Color = fillColor
    cardRg.Borders.LineStyle = xlContinuous
    cardRg.Borders.Weight = xlMedium
    With cardRg.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(246, 206, 62)
    End With

    ' Label area
    labelRg.Merge
    labelRg.Value = labelText
    With labelRg
        .Font.Bold = True
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ' Value area
    valueRg.Merge
    valueRg.Value = valueText
    With valueRg
        .Font.Bold = True
        .Font.Name = "Segoe UI"
        .Font.Size = 22
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    LogMsg "AddDashboardCard done: " & labelRange & "/" & valueRange
    Exit Sub

CleanFail:
    LogErr "AddDashboardCard"
End Sub

Private Sub StyleDashboardTable(ByVal rg As Range)
    Dim r As Long

    On Error Resume Next
    rg.Borders.LineStyle = xlContinuous
    rg.Borders.Weight = xlHairline
    For r = 1 To rg.Rows.Count
        If r Mod 2 = 0 Then rg.Rows(r).Interior.Color = RGB(242, 242, 242)
    Next r
End Sub

Private Sub StyleTableHeader(ByVal rg As Range)
    On Error Resume Next
    rg.Font.Bold = True
    rg.Interior.Color = RGB(31, 78, 121)
    rg.Font.Color = RGB(255, 255, 255)
    rg.HorizontalAlignment = xlCenter
End Sub

' ----------------------------------------------------------------------------
' Manager sheet formatting
' ----------------------------------------------------------------------------

Public Sub FormatManagerSheet(ByVal ws As Worksheet, ByVal managerName As String, ByVal pendingCount As Long, _
    ByVal managerLastCol As Long, ByVal managerPickupDateCol As Long, ByVal managerStartDateCol As Long, _
    ByVal managerAgeCol As Long, ByVal managerRepCol As Long, _
    ByVal firstRow As Long, ByVal lastRow As Long, ByVal ageCol As Long, ByVal pickupDateCol As Long, _
    ByVal currentData As Variant, ByVal managerRowsCol As Collection, ByVal hasPrior As Boolean, ByVal sourcePickupDateCol As Long)

    Dim r As Long, v As Variant
    Dim rg As Range
    Dim cardLabel As String, cardValue As String
    Dim m30 As Long, m60 As Long, m90 As Long, mMiss As Long
    Dim avgAge As Variant, pct30 As Variant
    Dim reportDateStr As String

    On Error GoTo CleanFail

    LogMsg "FormatManagerSheet start: " & managerName & " rows=" & firstRow & "-" & lastRow & " cols=1-" & managerLastCol

    reportDateStr = Format(Date, "mm/dd/yyyy")

    ' Calculate manager-specific stats (use source data column index)
    CalculateAgingStats currentData, managerRowsCol, sourcePickupDateCol, m30, m60, m90, mMiss
    avgAge = CalculateAverageAge(currentData, managerRowsCol, sourcePickupDateCol)
    If managerRowsCol.Count > 0 Then
        pct30 = m30 / managerRowsCol.Count
    Else
        pct30 = 0
    End If

    ' --- Manager header A1:H2 ---
    Set rg = ws.Range("A1:H2")
    rg.Merge
    rg.Value = managerName & "   |   as of " & reportDateStr
    With rg
        .Font.Bold = True
        .Font.Size = 18
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Rows("1:2").RowHeight = 27.95

    ' --- Manager KPI cards rows 4-10 ---
    ws.Rows("3:3").RowHeight = 16.5
    ws.Rows("4:6").RowHeight = 38.1
    ws.Rows("7:7").RowHeight = 15.75
    ws.Rows("8:10").RowHeight = 38.1

    cardLabel = "TOTAL PENDING PICKUPS" & vbCrLf & "Current"
    cardValue = CStr(managerRowsCol.Count)
    AddDashboardCard ws, "A4:B5", "A6:B6", RGB(221, 235, 247), cardLabel, cardValue

    cardLabel = "NEW SINCE PRIOR" & vbCrLf & "Since Prior Report"
    cardValue = IIf(hasPrior, "0", "N/A")
    AddDashboardCard ws, "C4:D5", "C6:D6", RGB(226, 239, 218), cardLabel, cardValue

    cardLabel = "AGING 30+ DAYS" & vbCrLf & "Older Pending Pickups"
    cardValue = CStr(m30)
    AddDashboardCard ws, "E4:F5", "E6:F6", RGB(244, 204, 204), cardLabel, cardValue

    cardLabel = "AGING 60+ DAYS" & vbCrLf & "Older Pending Pickups"
    cardValue = CStr(m60)
    AddDashboardCard ws, "G4:H5", "G6:H6", RGB(234, 153, 153), cardLabel, cardValue

    cardLabel = "AGING 90+ DAYS" & vbCrLf & "Older Pending Pickups"
    cardValue = CStr(m90)
    AddDashboardCard ws, "A8:B9", "A10:B10", RGB(224, 102, 102), cardLabel, cardValue

    cardLabel = "MISSING PICKUP DATE" & vbCrLf & "AUDIT: TP End Date"
    cardValue = CStr(mMiss)
    AddDashboardCard ws, "C8:D9", "C10:D10", RGB(255, 242, 204), cardLabel, cardValue

    cardLabel = "AVERAGE PICKUP AGE" & vbCrLf & "Pending Pickup Age"
    If IsNumeric(avgAge) Then
        cardValue = Format(CDbl(avgAge), "0.0")
    Else
        cardValue = "N/A"
    End If
    AddDashboardCard ws, "E8:F9", "E10:F10", RGB(209, 226, 242), cardLabel, cardValue

    cardLabel = "% OVER 30 DAYS" & vbCrLf & "Of Pending Pickups"
    cardValue = Format(pct30, "0.0%")
    AddDashboardCard ws, "G8:H9", "G10:H10", RGB(234, 209, 220), cardLabel, cardValue

    ' --- Data table formatting ---
    ' Write Pickup Age formulas into the age column for each data row
    For r = firstRow To lastRow
        If pickupDateCol > 0 Then
            Dim pdCell As String
            pdCell = ws.Cells(r, pickupDateCol).Address(False, False)
            ws.Cells(r, ageCol).Formula = "=IF(AND(ISNUMBER(" & pdCell & "), " & pdCell & ">=DATE(2000,1,1)), TODAY() - " & pdCell & ", """")"
        End If
    Next r

    ' Apply full data formatting (borders, font, alignment) to the entire table including header
    ApplyDataFormatting ws, MANAGER_HEADER_ROW, lastRow, managerLastCol

    ' Style header row
    Set rg = ws.Range(ws.Cells(MANAGER_HEADER_ROW, 1), ws.Cells(MANAGER_HEADER_ROW, managerLastCol))
    With rg
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Size = 11
    End With
    ws.Rows(MANAGER_HEADER_ROW).RowHeight = 36

    ' Apply rep banding (alternate colors by sales rep name)
    If managerRepCol > 0 Then
        ApplyRepBanding ws, firstRow, lastRow, managerLastCol, managerRepCol
    Else
        ' Fallback alternating row colors
        For r = firstRow To lastRow
            If (r - firstRow) Mod 2 = 0 Then
                ws.Range(ws.Cells(r, 1), ws.Cells(r, managerLastCol)).Interior.Color = RGB(242, 242, 242)
            End If
        Next r
    End If

    ' Apply age warning font color
    For r = firstRow To lastRow
        v = ws.Cells(r, ageCol).Value
        If IsNumeric(v) And Len(Trim$(CStr(v))) > 0 Then
            If CDbl(v) > 30 Then
                ws.Cells(r, ageCol).Font.Color = RGB(192, 0, 0)
            Else
                ws.Cells(r, ageCol).Font.Color = RGB(0, 0, 0)
            End If
        Else
            ws.Cells(r, ageCol).Font.Color = RGB(0, 0, 0)
        End If
    Next r

    ' Highlight missing pickup dates
    HighlightMissingPickupDates ws, firstRow, lastRow, pickupDateCol

    ' Format Yes/No columns
    FormatYesNoColumns ws, MANAGER_HEADER_ROW, lastRow, managerLastCol

    ' Auto-fit and set min widths
    ws.Range(ws.Cells(MANAGER_HEADER_ROW, 1), ws.Cells(LastUsedRow(ws), managerLastCol)).Columns.AutoFit
    SetManagerColumnWidths ws, managerLastCol

    ' Freeze panes below header
    ws.Activate
    On Error Resume Next
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitRow = MANAGER_HEADER_ROW
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 85
    On Error GoTo 0

    LogMsg "FormatManagerSheet done: " & managerName
    Exit Sub

CleanFail:
    LogErr "FormatManagerSheet"
End Sub

Private Sub ApplyRepBanding(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal lastCol As Long, ByVal repCol As Long)
    Dim r As Long, currentRep As String, prevRep As String
    Dim useAlt As Boolean

    On Error Resume Next
    If repCol <= 0 Or firstRow > lastRow Then Exit Sub

    useAlt = False
    prevRep = ""
    For r = firstRow To lastRow
        currentRep = LCase$(Trim$(CStr(ws.Cells(r, repCol).Value)))
        If r = firstRow Or currentRep <> prevRep Then
            If r > firstRow Then useAlt = Not useAlt
        End If
        If useAlt Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = RGB(217, 225, 242)
        Else
            ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = RGB(255, 255, 255)
        End If
        prevRep = currentRep
    Next r
End Sub

' ----------------------------------------------------------------------------
' Table / row formatting
' ----------------------------------------------------------------------------

Private Sub HighlightMissingPickupDates(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal pickupDateCol As Long)
    Dim r As Long, v As Variant, dt As Date

    On Error Resume Next
    If pickupDateCol <= 0 Or firstRow > lastRow Then Exit Sub

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
End Sub

Private Sub FormatYesNoColumns(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long, ByVal lastCol As Long)
    Dim c As Long, r As Long, h As String, v As String

    On Error Resume Next
    If lastRow <= headerRow Then Exit Sub

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
End Sub

Private Sub SetManagerColumnWidths(ByVal ws As Worksheet, ByVal lastCol As Long)
    Dim c As Long, headerText As String, minWidth As Double

    If lastCol <= 0 Then Exit Sub

    On Error Resume Next
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

Public Sub FormatRawData(ByVal ws As Worksheet, ByVal lastCol As Long, ByVal lastRow As Long)
    ApplyDataFormatting ws, 1, lastRow, lastCol
    ws.Tab.Color = RGB(128, 128, 128)

    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "1:1"
    End With
    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 85
    On Error GoTo 0
End Sub

Private Sub ApplyDataFormatting(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long, ByVal lastCol As Long)
    Dim rg As Range

    On Error Resume Next
    Set rg = ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, lastCol))
    rg.Font.Name = "Segoe UI"
    rg.Borders.LineStyle = xlContinuous
    rg.HorizontalAlignment = xlCenter
    With rg.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(246, 206, 62)
    End With
End Sub

' ----------------------------------------------------------------------------
' Email snapshot helper
' ----------------------------------------------------------------------------

Public Function CopySheetSnapshot(ByVal ws As Worksheet, ByVal rng As Range, ByVal filePath As String) As Boolean
    Dim cht As ChartObject
    On Error Resume Next
    CopySheetSnapshot = False

    If ws Is Nothing Or rng Is Nothing Then Exit Function

    rng.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    If Err.Number <> 0 Then Exit Function

    Set cht = ws.ChartObjects.Add(Left:=0, Top:=0, Width:=rng.Width, Height:=rng.Height)
    If Err.Number <> 0 Then Exit Function

    cht.Chart.Paste
    cht.Chart.Export filePath, "PNG"
    cht.Delete

    CopySheetSnapshot = (Dir$(filePath) <> "")
End Function
