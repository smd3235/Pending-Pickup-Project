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

    Dim rg As Range, labelRg As Range, valueRow As Range
    Dim fillColor As Long
    Dim labelText As String, subText As String, valueText As String
    Dim r As Long, c As Long
    Dim mgrName As String
    Dim i As Long

    ' Title block
    Set rg = ws.Range("A1:D1")
    rg.Merge
    rg.Value = "PENDING PICKUP REPORT DASHBOARD"
    rg.Font.Bold = True
    rg.Font.Size = 18
    rg.Font.Color = RGB(255, 255, 255)
    rg.Interior.Color = RGB(31, 78, 121)
    rg.HorizontalAlignment = xlCenter
    rg.VerticalAlignment = xlCenter
    rg.RowHeight = 30

    ' Subtitle
    Set rg = ws.Range("A2:D2")
    rg.Merge
    rg.Value = "Generated: " & reportStamp & "  |  Source: " & sourceName
    rg.Font.Size = 10
    rg.Font.Italic = True
    rg.Interior.Color = RGB(217, 234, 247)
    rg.HorizontalAlignment = xlCenter

    ' KPI cards row 4
    fillColor = RGB(217, 234, 247)

    ' Total Pending
    labelText = "Total Pending Records"
    valueText = CStr(currentCount)
    Call AddDashboardCard(ws, "B4", "D6", fillColor, labelText, valueText)

    ' New since prior
    labelText = "New Since Prior Report"
    valueText = CStr(newCount)
    Call AddDashboardCard(ws, "F4", "H6", fillColor, labelText, valueText)

    ' Still pending
    labelText = "Still Pending"
    valueText = CStr(stillCount)
    Call AddDashboardCard(ws, "J4", "L6", fillColor, labelText, valueText)

    ' Cleared since prior
    labelText = "Cleared Since Prior"
    valueText = CStr(clearedCount)
    Call AddDashboardCard(ws, "N4", "P6", fillColor, labelText, valueText)

    ' Aging stats row 8
    labelText = "Aging 30+ Days"
    valueText = CStr(age30Count)
    Call AddDashboardCard(ws, "B8", "D10", fillColor, labelText, valueText)

    labelText = "Aging 60+ Days"
    valueText = CStr(age60Count)
    Call AddDashboardCard(ws, "F8", "H10", fillColor, labelText, valueText)

    labelText = "Aging 90+ Days"
    valueText = CStr(age90Count)
    Call AddDashboardCard(ws, "J8", "L10", fillColor, labelText, valueText)

    labelText = "Missing Pickup Dates"
    valueText = CStr(missingDateCount)
    Call AddDashboardCard(ws, "N8", "P10", fillColor, labelText, valueText)

    ' Manager table header
    Set rg = ws.Range("A13")
    rg.Value = "MANAGER BREAKDOWN"
    rg.Font.Bold = True
    rg.Font.Size = 14
    rg.Font.Color = RGB(255, 255, 255)
    rg.Interior.Color = RGB(31, 78, 121)

    ' Manager table headers
    Dim headers As Variant
    headers = Array("Manager", "Pending Count", "30+ Days", "60+ Days", "90+ Days", "Missing Dates")
    For c = 0 To UBound(headers)
        Set rg = ws.Cells(14, c + 1)
        rg.Value = headers(c)
        rg.Font.Bold = True
        rg.Interior.Color = RGB(217, 234, 247)
        rg.HorizontalAlignment = xlCenter
    Next c

    ' Manager table rows
    r = 15
    For i = LBound(managers) To UBound(managers)
        mgrName = CStr(managers(i))
        ws.Cells(r, 1).Value = mgrName
        ws.Cells(r, 2).Value = managerRows(mgrName).Count
        ws.Cells(r, 3).Value = 0
        ws.Cells(r, 4).Value = 0
        ws.Cells(r, 5).Value = 0
        ws.Cells(r, 6).Value = 0
        r = r + 1
    Next i

    StyleTableHeader ws.Range("A14:F14")
    StyleDashboardTable ws.Range("A15:F" & (r - 1))
    ws.Columns("A:P").AutoFit
End Sub

Private Sub AddDashboardCard(ByVal ws As Worksheet, ByVal topLeft As String, ByVal bottomRight As String, _
    ByVal fillColor As Long, ByVal labelText As String, ByVal valueText As String)

    Dim rg As Range, labelRg As Range, valueRow As Range
    Dim rows As Long, cols As Long

    Set rg = ws.Range(topLeft & ":" & bottomRight)
    rows = rg.Rows.Count
    cols = rg.Columns.Count

    ' Color the whole card range (no parent merge - avoids nested merge conflicts)
    rg.Interior.Color = fillColor
    rg.Borders.LineStyle = xlContinuous
    rg.Borders.Weight = xlMedium

    ' Label occupies top rows
    Set labelRg = rg.Resize(rows - 1, cols)
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

    ' Value gets its own merged bottom row
    Set valueRow = rg.Offset(rows - 1, 0).Resize(1, cols)
    valueRow.Merge
    valueRow.Value = valueText
    With valueRow
        .Font.Bold = True
        .Font.Name = "Segoe UI"
        .Font.Size = 22
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
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
    ByVal firstRow As Long, ByVal lastRow As Long, ByVal ageCol As Long, ByVal pickupDateCol As Long)

    Dim r As Long, v As Variant

    On Error Resume Next

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
