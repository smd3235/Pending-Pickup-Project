Attribute VB_Name = "modDialog"
Option Explicit

' ============================================================================
'  Module: modDialog
'  File/folder dialogs, folder settings persistence, prior-report discovery,
'  output path building, and macro-enabled workbook enforcement.
' ============================================================================

' ----------------------------------------------------------------------------
' File picker
' ----------------------------------------------------------------------------

Public Function SelectExcelFile(ByVal title As String, ByVal suggestedFolder As String, _
    ByVal suggestedFile As String, ByVal allowCancel As Boolean) As Variant

    Dim fDialog As Object
    Dim result As Long
    Dim initialFolder As String

    On Error Resume Next
    SelectExcelFile = False

    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    If fDialog Is Nothing Then Exit Function

    With fDialog
        .Title = title
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb", 1
        .Filters.Add "All Files", "*.*", 2
        .FilterIndex = 1

        ' Set initial folder
        initialFolder = suggestedFolder
        If Len(Trim$(initialFolder)) = 0 Then initialFolder = ThisWorkbook.Path
        If Dir$(initialFolder, vbDirectory) <> "" Then
            .InitialFileName = initialFolder & "\"
        End If

        ' Set initial file if suggested
        If Len(Trim$(suggestedFile)) > 0 Then
            .InitialFileName = suggestedFile
        End If

        result = .Show
        If result = -1 Then
            SelectExcelFile = .SelectedItems(1)
        Else
            SelectExcelFile = False
        End If
    End With
End Function

' ----------------------------------------------------------------------------
' Folder settings (registry)
' ----------------------------------------------------------------------------

Public Function GetSavedFolder(ByVal keyName As String) As String
    On Error Resume Next
    GetSavedFolder = GetSetting("PendingPickupReport", "Folders", keyName, "")
End Function

Public Sub SaveFolderSetting(ByVal keyName As String, ByVal folderPath As String)
    On Error Resume Next
    SaveSetting "PendingPickupReport", "Folders", keyName, folderPath
End Sub

' ----------------------------------------------------------------------------
' Prior report discovery
' ----------------------------------------------------------------------------

Public Function FindMostRecentPendingPickupReport(ByVal folderPath As String, ByVal excludeFile As String) As String
    Dim fso As Object, folder As Object, file As Object
    Dim mostRecent As String, mostRecentDate As Date
    Dim fileName As String, fileDate As Date

    On Error Resume Next
    FindMostRecentPendingPickupReport = ""

    If Len(Trim$(folderPath)) = 0 Then Exit Function
    If Dir$(folderPath, vbDirectory) = "" Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    mostRecentDate = 0
    For Each file In folder.Files
        fileName = LCase$(file.Name)
        If (InStr(1, fileName, "pending pickup", vbTextCompare) > 0 Or InStr(1, fileName, "pendingpickup", vbTextCompare) > 0) Then
            If InStr(1, fileName, ".xlsx") > 0 Or InStr(1, fileName, ".xlsm") > 0 Or InStr(1, fileName, ".xls") > 0 Then
                If StrComp(file.Path, excludeFile, vbTextCompare) <> 0 Then
                    fileDate = file.DateLastModified
                    If fileDate > mostRecentDate Then
                        mostRecentDate = fileDate
                        mostRecent = file.Path
                    End If
                End If
            End If
        End If
    Next file

    FindMostRecentPendingPickupReport = mostRecent
End Function

' ----------------------------------------------------------------------------
' Output path
' ----------------------------------------------------------------------------

Public Function BuildOutputPath(ByVal masterWb As Workbook) As String
    Dim reportsDir As String, stamp As String, outName As String, outPath As String
    Dim fso As Object

    On Error Resume Next
    reportsDir = masterWb.Path & "\Reports"
    If Len(Dir$(reportsDir, vbDirectory)) = 0 Then MkDir reportsDir

    stamp = Format(Now, "yyyy-mm-dd_HHMMSS")
    outName = "Pending Pickup " & stamp & ".xlsx"
    outPath = reportsDir & "\" & outName

    BuildOutputPath = outPath
End Function

' ----------------------------------------------------------------------------
' Macro-enabled workbook enforcement
' ----------------------------------------------------------------------------

Public Function EnsureMacroEnabledWorkbook() As Boolean
    Dim f As String
    On Error Resume Next
    EnsureMacroEnabledWorkbook = False

    f = ThisWorkbook.FullName
    If InStr(1, f, ".xlsm", vbTextCompare) > 0 Or InStr(1, f, ".xlsb", vbTextCompare) > 0 Or InStr(1, f, ".xltm", vbTextCompare) > 0 Then
        EnsureMacroEnabledWorkbook = True
        Exit Function
    End If

    MsgBox "This workbook is saved as a non-macro-enabled file (" & f & ")." & vbCrLf & vbCrLf & _
           "Please save it as an Excel Macro-Enabled Workbook (.xlsm) before running the report.", vbExclamation
End Function

' ----------------------------------------------------------------------------
' Email config (manager emails from hidden template tab; director from hidden template tab)
' ----------------------------------------------------------------------------

Private Function ReadTemplateConfig(ByVal sheetName As String, ByVal cellAddr As String, ByVal fallbackSheet As String, ByVal fallbackCell As String, ByVal logLabel As String) As String
    On Error Resume Next
    Dim v As String
    v = Trim$(CStr(ThisWorkbook.Worksheets(sheetName).Range(cellAddr).Value))
    If Err.Number <> 0 Or Len(v) = 0 Then
        Err.Clear
        v = Trim$(CStr(ThisWorkbook.Worksheets(fallbackSheet).Range(fallbackCell).Value))
    End If
    If Err.Number <> 0 Then v = ""
    Err.Clear
    LogMsg "Email config read: " & logLabel & "=" & Left$(Replace(v, vbCr, " "), 80)
    ReadTemplateConfig = v
End Function

Private Function ReadTemplateBody(ByVal sheetName As String, ByVal cellAddr As String, ByVal fallbackSheet As String, ByVal fallbackCell As String, ByVal defaultBody As String, ByVal logLabel As String) As String
    On Error Resume Next
    Dim v As String
    v = CStr(ThisWorkbook.Worksheets(sheetName).Range(cellAddr).Value)
    If Err.Number <> 0 Or Len(Trim$(v)) = 0 Then
        Err.Clear
        v = CStr(ThisWorkbook.Worksheets(fallbackSheet).Range(fallbackCell).Value)
    End If
    If Err.Number <> 0 Or Len(Trim$(v)) = 0 Then v = defaultBody
    Err.Clear
    LogMsg "Email config read: " & logLabel & "=" & Left$(Replace(v, vbCr, " "), 80)
    ReadTemplateBody = v
End Function

Public Function GetEmailTo() As String
    On Error Resume Next
    GetEmailTo = Trim$(CStr(ThisWorkbook.Worksheets(EMAIL_CONFIG_SHEET).Range(EMAIL_CONFIG_TO_CELL).Value))
    If Err.Number <> 0 Then GetEmailTo = ""
    Err.Clear
    LogMsg "Email config read: Manager To=" & GetEmailTo
End Function

Public Function GetEmailCC() As String
    GetEmailCC = ReadTemplateConfig(EMAIL_CONFIG_SHEET, EMAIL_CONFIG_CC_CELL, LEGACY_EMAIL_CONFIG_SHEET, LEGACY_EMAIL_CONFIG_CC_CELL, "Manager CC")
End Function

Public Function GetEmailSubject(ByVal defaultSubject As String) As String
    Dim v As String
    v = ReadTemplateConfig(EMAIL_CONFIG_SHEET, EMAIL_CONFIG_SUBJECT_CELL, LEGACY_EMAIL_CONFIG_SHEET, LEGACY_EMAIL_CONFIG_SUBJECT_CELL, "Manager Subject")
    If Len(v) = 0 Then v = defaultSubject
    GetEmailSubject = v
End Function

Public Function GetEmailBody(ByVal defaultBody As String) As String
    GetEmailBody = ReadTemplateBody(EMAIL_CONFIG_SHEET, EMAIL_CONFIG_BODY_CELL, LEGACY_EMAIL_CONFIG_SHEET, LEGACY_EMAIL_CONFIG_BODY_CELL, defaultBody, "Manager Body")
End Function

Public Function GetDirectorEmailTo() As String
    On Error Resume Next
    GetDirectorEmailTo = Trim$(CStr(ThisWorkbook.Worksheets(DIRECTOR_EMAIL_SHEET).Range(DIRECTOR_EMAIL_TO_CELL).Value))
    If Err.Number <> 0 Then GetDirectorEmailTo = ""
    Err.Clear
    LogMsg "Email config read: Director To=" & GetDirectorEmailTo
End Function

Public Function GetDirectorEmailCC() As String
    On Error Resume Next
    GetDirectorEmailCC = Trim$(CStr(ThisWorkbook.Worksheets(DIRECTOR_EMAIL_SHEET).Range(DIRECTOR_EMAIL_CC_CELL).Value))
    If Err.Number <> 0 Then GetDirectorEmailCC = ""
    Err.Clear
    LogMsg "Email config read: Director CC=" & GetDirectorEmailCC
End Function

Public Function GetDirectorEmailSubject(ByVal defaultSubject As String) As String
    On Error Resume Next
    GetDirectorEmailSubject = Trim$(CStr(ThisWorkbook.Worksheets(DIRECTOR_EMAIL_SHEET).Range(DIRECTOR_EMAIL_SUBJECT_CELL).Value))
    If Err.Number <> 0 Or Len(GetDirectorEmailSubject) = 0 Then GetDirectorEmailSubject = defaultSubject
    Err.Clear
    LogMsg "Email config read: Director Subject=" & GetDirectorEmailSubject
End Function

Public Function GetDirectorEmailBody(ByVal defaultBody As String) As String
    On Error Resume Next
    GetDirectorEmailBody = CStr(ThisWorkbook.Worksheets(DIRECTOR_EMAIL_SHEET).Range(DIRECTOR_EMAIL_BODY_CELL).Value)
    If Err.Number <> 0 Or Len(Trim$(GetDirectorEmailBody)) = 0 Then GetDirectorEmailBody = defaultBody
    Err.Clear
    LogMsg "Email config read: Director Body=" & Left$(Replace(GetDirectorEmailBody, vbCr, " "), 80)
End Function

Public Sub SetupEmailCC()
    Dim ws As Worksheet
    On Error GoTo CleanFail
    Set ws = ThisWorkbook.Worksheets(EMAIL_CONFIG_SHEET)
    ws.Visible = xlSheetVisible
    ws.Activate
    ws.Range(EMAIL_CONFIG_TO_CELL).Select
    LogMsg "Email setup: user opened Manager Email Template sheet"
    MsgBox "Edit manager email settings on '" & EMAIL_CONFIG_SHEET & "':" & vbCrLf & _
           "  TO - " & EMAIL_CONFIG_TO_CELL & " (separate multiple addresses with ;)" & vbCrLf & _
           "  CC - " & EMAIL_CONFIG_CC_CELL & " (separate multiple addresses with ;)" & vbCrLf & _
           "  SUBJECT - " & EMAIL_CONFIG_SUBJECT_CELL & " (leave blank for default)" & vbCrLf & _
           "  BODY - " & EMAIL_CONFIG_BODY_CELL & " (leave blank for default intro)" & vbCrLf & vbCrLf & _
           "Use the 'Return to START HERE' button on the sheet to go back and re-hide this tab." & vbCrLf & _
           "Director/dashboard email settings are on '" & DIRECTOR_EMAIL_SHEET & "'.", vbInformation
    Exit Sub
CleanFail:
    LogErr "SetupEmailCC"
    MsgBox "Unable to open '" & EMAIL_CONFIG_SHEET & "'.", vbExclamation
End Sub

Public Sub SetupDirectorEmail()
    Dim ws As Worksheet
    On Error GoTo CleanFail
    Set ws = ThisWorkbook.Worksheets(DIRECTOR_EMAIL_SHEET)
    ws.Visible = xlSheetVisible
    ws.Activate
    ws.Range(DIRECTOR_EMAIL_TO_CELL).Select
    LogMsg "Email setup: user opened Director email template sheet"
    MsgBox "Edit dashboard/director email settings on '" & DIRECTOR_EMAIL_SHEET & "':" & vbCrLf & _
           "  TO - " & DIRECTOR_EMAIL_TO_CELL & " (separate multiple addresses with ;)" & vbCrLf & _
           "  CC - " & DIRECTOR_EMAIL_CC_CELL & " (separate multiple addresses with ;)" & vbCrLf & _
           "  SUBJECT - " & DIRECTOR_EMAIL_SUBJECT_CELL & " (leave blank for default)" & vbCrLf & _
           "  BODY - " & DIRECTOR_EMAIL_BODY_CELL & " (leave blank for default intro)" & vbCrLf & vbCrLf & _
           "Use the 'Return to START HERE' button on the sheet to go back and re-hide this tab.", vbInformation
    Exit Sub
CleanFail:
    LogErr "SetupDirectorEmail"
    MsgBox "Unable to open '" & DIRECTOR_EMAIL_SHEET & "'.", vbExclamation
End Sub

Public Sub ReturnToStartHere()
    Dim callerName As String
    Dim shp As Shape
    Dim wsReturn As Worksheet
    Dim wsStart As Worksheet

    On Error GoTo CleanFail
    callerName = CStr(Application.Caller)
    Set wsStart = ThisWorkbook.Worksheets("START HERE")

    Set wsReturn = Nothing
    On Error Resume Next
    Set shp = ActiveSheet.Shapes(callerName)
    If Not shp Is Nothing Then Set wsReturn = shp.Parent
    On Error GoTo CleanFail

    If wsReturn Is Nothing Then
        ' Fallback: if called from an email template sheet, return/hide it
        If ActiveSheet.Name = EMAIL_CONFIG_SHEET Or ActiveSheet.Name = DIRECTOR_EMAIL_SHEET Then
            Set wsReturn = ActiveSheet
        End If
    End If

    wsStart.Visible = xlSheetVisible
    wsStart.Activate

    If Not wsReturn Is Nothing Then
        If wsReturn.Name <> wsStart.Name Then wsReturn.Visible = xlSheetHidden
    End If

    LogMsg "Email template: returned to START HERE and hid sheet"
    Exit Sub

CleanFail:
    LogErr "ReturnToStartHere"
    On Error Resume Next
    ThisWorkbook.Worksheets("START HERE").Visible = xlSheetVisible
    ThisWorkbook.Worksheets("START HERE").Activate
End Sub

' ----------------------------------------------------------------------------
' Helper: wait for Outlook Word editor to be available after opening inspector
' ----------------------------------------------------------------------------
Private Function GetWordEditor(ByVal mail As Object, Optional ByVal maxWaitSeconds As Single = 8) As Object
    On Error Resume Next
    GetWordEditor = Nothing
    Dim startTime As Single
    startTime = Timer
    Do
        Dim inspector As Object
        Set inspector = mail.GetInspector
        If Not inspector Is Nothing Then
            Dim wd As Object
            Set wd = inspector.WordEditor
            If Not wd Is Nothing Then
                Set GetWordEditor = wd
                Exit Function
            End If
        End If
        DoEvents
    Loop While Timer - startTime < maxWaitSeconds
End Function

Private Function GetDashboardSnapshotRange(ByVal ws As Worksheet) As Range
    On Error Resume Next
    Dim lastRow As Long, lastCol As Long
    lastRow = LastUsedRow(ws)
    lastCol = LastUsedColumn(ws)

    Dim co As ChartObject
    For Each co In ws.ChartObjects
        If co.BottomRightCell.Row > lastRow Then lastRow = co.BottomRightCell.Row
        If co.BottomRightCell.Column > lastCol Then lastCol = co.BottomRightCell.Column
    Next co

    If lastRow < 1 Then lastRow = 25
    If lastCol < 1 Then lastCol = 16 ' Column P
    If lastCol < 16 Then lastCol = 16

    Set GetDashboardSnapshotRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
End Function

Private Function GetManagerSnapshotRange(ByVal ws As Worksheet) As Range
    On Error Resume Next
    Dim lastRow As Long, lastCol As Long
    lastRow = LastUsedRow(ws)

    ' Prefer the data table header row bounds to avoid copying email button columns.
    lastCol = LastUsedColumn(ws)
    Dim headerLast As Long
    headerLast = ws.Cells(MANAGER_HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    If headerLast > 0 And headerLast < lastCol Then lastCol = headerLast

    If lastRow < MANAGER_HEADER_ROW Then lastRow = MANAGER_HEADER_ROW
    If lastCol < 1 Then lastCol = 1

    Set GetManagerSnapshotRange = ws.Range(ws.Cells(MANAGER_HEADER_ROW, 1), ws.Cells(lastRow, lastCol))
End Function

Private Sub NotifyEmailFailure(ByVal context As String, ByVal userMessage As String)
    LogMsg "Email FAILED [" & context & "]: " & userMessage
    MsgBox userMessage & vbCrLf & vbCrLf & "See macro-debug.log for details.", vbExclamation, "Email Not Created"
End Sub

' ----------------------------------------------------------------------------
' Send email to a manager using copy/paste of Excel range into Outlook Word editor
' ----------------------------------------------------------------------------

Public Sub SendManagerEmail(ByVal ws As Worksheet, ByVal managerName As String, ByVal rowCount As Long)
    On Error GoTo CleanFail

    Dim snapRg As Range
    Set snapRg = GetManagerSnapshotRange(ws)

    LogMsg "Email: preparing Excel content for manager " & managerName & " (rows=" & rowCount & ", range=" & snapRg.Address(False, False) & ")"

    If ws Is Nothing Or snapRg Is Nothing Then
        NotifyEmailFailure "SendManagerEmail", "Could not determine the manager report range to copy."
        Exit Sub
    End If

    ' Copy the Excel range to the clipboard so formatting and spacing are preserved when pasted into Outlook.
    snapRg.Copy
    If Err.Number <> 0 Then
        NotifyEmailFailure "SendManagerEmail", "Could not copy the manager report range to the clipboard."
        Err.Clear
        GoTo CleanExit
    End If

    Dim outlook As Object, mail As Object
    Set outlook = CreateObject("Outlook.Application")
    If outlook Is Nothing Then
        NotifyEmailFailure "SendManagerEmail", "Outlook is not available. Open Outlook and try again."
        GoTo CleanExit
    End If
    LogMsg "Email: Outlook application created for manager " & managerName
    Set mail = outlook.CreateItem(0)

    With mail
        .To = GetEmailTo()
        .Subject = GetEmailSubject(EMAIL_SUBJECT_MANAGER & " - " & managerName & " (" & rowCount & " pending)")
        .CC = GetEmailCC()

        ' Display the inspector so we can access the Word editor and paste formatted Excel content.
        .Display
        DoEvents

        Dim wordDoc As Object
        Set wordDoc = GetWordEditor(mail, 8)
        If wordDoc Is Nothing Then
            NotifyEmailFailure "SendManagerEmail", "Outlook opened, but the email editor was not ready. Try again in a few seconds."
            GoTo CleanExit
        End If

        Dim introText As String
        introText = GetEmailBody("Please review the Pending Pickup report below. You have " & rowCount & " records requiring pickup action.")

        ' Insert intro text at the top.
        Dim bodyRange As Object
        Set bodyRange = wordDoc.Range(0, 0)
        bodyRange.Text = "Hi " & managerName & "," & vbCrLf & introText & vbCrLf & vbCrLf

        ' Move to the end and paste the copied Excel range (preserves formatting).
        Set bodyRange = wordDoc.Range(wordDoc.Characters.Count, wordDoc.Characters.Count)
        bodyRange.Paste

        ' Add closing text.
        Set bodyRange = wordDoc.Range(wordDoc.Characters.Count, wordDoc.Characters.Count)
        bodyRange.Text = vbCrLf & vbCrLf & "Please review and respond as needed." & vbCrLf & "Pending Pickup Report System"

        LogMsg "Email: inspector ready for manager " & managerName & " (To=" & .To & ", CC=" & .CC & ", Subject=" & .Subject & ")"
    End With

    LogMsg "Email: inspector displayed for manager " & managerName
    GoTo CleanExit

CleanFail:
    LogErr "SendManagerEmail"
    MsgBox "An unexpected error occurred while creating the manager email." & vbCrLf & vbCrLf & "See macro-debug.log for details.", vbExclamation, "Email Not Created"

CleanExit:
    On Error Resume Next
    Application.CutCopyMode = False
End Sub

' ----------------------------------------------------------------------------
' Send dashboard email to Director using copy/paste of Excel range into Outlook Word editor
' ----------------------------------------------------------------------------

Public Sub SendDashboardEmail(ByVal ws As Worksheet, ByVal rowCount As Long)
    On Error GoTo CleanFail

    Dim snapRg As Range
    Set snapRg = GetDashboardSnapshotRange(ws)

    LogMsg "Email: preparing Excel content for dashboard (rows=" & rowCount & ", range=" & snapRg.Address(False, False) & ")"

    If ws Is Nothing Or snapRg Is Nothing Then
        NotifyEmailFailure "SendDashboardEmail", "Could not determine the dashboard range to copy."
        Exit Sub
    End If

    ' Copy the Excel range to the clipboard so formatting and spacing are preserved when pasted into Outlook.
    snapRg.Copy
    If Err.Number <> 0 Then
        NotifyEmailFailure "SendDashboardEmail", "Could not copy the dashboard range to the clipboard."
        Err.Clear
        GoTo CleanExit
    End If

    Dim outlook As Object, mail As Object
    Set outlook = CreateObject("Outlook.Application")
    If outlook Is Nothing Then
        NotifyEmailFailure "SendDashboardEmail", "Outlook is not available. Open Outlook and try again."
        GoTo CleanExit
    End If
    LogMsg "Email: Outlook application created for dashboard email"
    Set mail = outlook.CreateItem(0)

    With mail
        .To = GetDirectorEmailTo()
        .Subject = GetDirectorEmailSubject(EMAIL_SUBJECT_DIRECTOR & " (" & Format(Date, "m/d/yyyy") & ")")
        .CC = GetDirectorEmailCC()

        ' Display the inspector so we can access the Word editor and paste formatted Excel content.
        .Display
        DoEvents

        Dim wordDoc As Object
        Set wordDoc = GetWordEditor(mail, 8)
        If wordDoc Is Nothing Then
            NotifyEmailFailure "SendDashboardEmail", "Outlook opened, but the email editor was not ready. Try again in a few seconds."
            GoTo CleanExit
        End If

        Dim introText As String
        If rowCount > 0 Then
            introText = GetDirectorEmailBody("Please find the Pending Pickup daily dashboard below. Current pending records: " & rowCount & ".")
        Else
            introText = GetDirectorEmailBody("Please find the Pending Pickup daily dashboard below.")
        End If

        ' Insert intro text at the top.
        Dim bodyRange As Object
        Set bodyRange = wordDoc.Range(0, 0)
        bodyRange.Text = "Hi Director," & vbCrLf & introText & vbCrLf & vbCrLf

        ' Move to the end and paste the copied Excel range (preserves formatting).
        Set bodyRange = wordDoc.Range(wordDoc.Characters.Count, wordDoc.Characters.Count)
        bodyRange.Paste

        ' Add closing text.
        Set bodyRange = wordDoc.Range(wordDoc.Characters.Count, wordDoc.Characters.Count)
        bodyRange.Text = vbCrLf & vbCrLf & "Pending Pickup Report System"

        LogMsg "Email: inspector ready for dashboard (To=" & .To & ", CC=" & .CC & ", Subject=" & .Subject & ")"
    End With

    LogMsg "Email: inspector displayed for dashboard"
    GoTo CleanExit

CleanFail:
    LogErr "SendDashboardEmail"
    MsgBox "An unexpected error occurred while creating the dashboard email." & vbCrLf & vbCrLf & "See macro-debug.log for details.", vbExclamation, "Email Not Created"

CleanExit:
    On Error Resume Next
    Application.CutCopyMode = False
End Sub

' ----------------------------------------------------------------------------
' Button click handlers - use Application.Caller + AlternativeText to identify target sheet
' ----------------------------------------------------------------------------

Private Function ResolveCallerShape(ByVal callerName As String) As Shape
    On Error Resume Next
    Set ResolveCallerShape = Nothing
    If Len(callerName) = 0 Then Exit Function

    Set ResolveCallerShape = ActiveSheet.Shapes(callerName)
    If ResolveCallerShape Is Nothing Then
        Set ResolveCallerShape = ThisWorkbook.ActiveSheet.Shapes(callerName)
    End If
    If ResolveCallerShape Is Nothing Then
        Dim ws As Worksheet
        For Each ws In ThisWorkbook.Worksheets
            Set ResolveCallerShape = ws.Shapes(callerName)
            If Not ResolveCallerShape Is Nothing Then Exit Function
        Next ws
    End If
End Function

Public Sub EmailSelectedManagerTab()
    On Error GoTo CleanFail
    LogMsg "Email: user clicked manager tab email button"

    Dim callerName As String
    callerName = CStr(Application.Caller)
    If Len(callerName) = 0 Then
        NotifyEmailFailure "EmailSelectedManagerTab", "Could not identify which manager email button was clicked."
        Exit Sub
    End If

    Dim shp As Shape
    Set shp = ResolveCallerShape(callerName)
    If shp Is Nothing Then
        NotifyEmailFailure "EmailSelectedManagerTab", "Could not locate the clicked email button."
        Exit Sub
    End If

    Dim targetWb As Workbook
    Set targetWb = shp.Parent.Parent

    Dim targetSheetName As String
    targetSheetName = Trim$(CStr(shp.AlternativeText))
    If Len(targetSheetName) = 0 Then
        NotifyEmailFailure "EmailSelectedManagerTab", "The clicked email button is not linked to a manager sheet."
        Exit Sub
    End If

    Dim targetSheet As Worksheet
    On Error Resume Next
    Set targetSheet = targetWb.Worksheets(targetSheetName)
    On Error GoTo CleanFail
    If targetSheet Is Nothing Then
        NotifyEmailFailure "EmailSelectedManagerTab", "The target manager sheet '" & targetSheetName & "' could not be found."
        Exit Sub
    End If

    Dim mgrName As String
    mgrName = Trim$(Replace(shp.TextFrame2.TextRange.Text, "Email ", ""))
    If Len(mgrName) = 0 Then mgrName = targetSheetName

    Dim rowCount As Long
    rowCount = LastUsedRow(targetSheet) - MANAGER_HEADER_ROW
    If rowCount < 0 Then rowCount = 0

    SendManagerEmail targetSheet, mgrName, rowCount
    Exit Sub

CleanFail:
    LogErr "EmailSelectedManagerTab"
    MsgBox "An unexpected error occurred while handling the manager email button." & vbCrLf & vbCrLf & "See macro-debug.log for details.", vbExclamation, "Email Not Created"
End Sub

Public Sub EmailDashboard()
    On Error GoTo CleanFail
    LogMsg "Email: user clicked dashboard email button"

    Dim callerName As String
    callerName = CStr(Application.Caller)

    Dim shp As Shape
    Set shp = ResolveCallerShape(callerName)

    Dim targetWb As Workbook
    If Not shp Is Nothing Then
        Set targetWb = shp.Parent.Parent
    Else
        Set targetWb = ThisWorkbook
    End If

    Dim targetSheet As Worksheet
    If Not shp Is Nothing And Len(Trim$(CStr(shp.AlternativeText))) > 0 Then
        On Error Resume Next
        Set targetSheet = targetWb.Worksheets(Trim$(CStr(shp.AlternativeText)))
        On Error GoTo CleanFail
    End If
    If targetSheet Is Nothing Then
        On Error Resume Next
        Set targetSheet = targetWb.Worksheets("Dashboard")
        On Error GoTo CleanFail
    End If
    If targetSheet Is Nothing Then
        NotifyEmailFailure "EmailDashboard", "Could not find the Dashboard worksheet."
        Exit Sub
    End If

    ' Count pending records from the manager breakdown table when available; fallback to last used row.
    Dim rowCount As Long
    rowCount = 0
    On Error Resume Next
    Dim breakdownHeader As Range
    Set breakdownHeader = targetSheet.Cells.Find(What:="Assigned Sales Manager", LookIn:=xlValues, LookAt:=xlWhole)
    If Not breakdownHeader Is Nothing Then
        Dim r As Long
        For r = breakdownHeader.Row + 1 To LastUsedRow(targetSheet)
            If Len(Trim$(CStr(targetSheet.Cells(r, breakdownHeader.Column).Value))) = 0 Then Exit For
            If StrComp(Trim$(CStr(targetSheet.Cells(r, breakdownHeader.Column).Value)), "TOTAL", vbTextCompare) = 0 Then Exit For
            rowCount = rowCount + 1
        Next r
    End If
    On Error GoTo CleanFail
    If rowCount = 0 Then rowCount = LastUsedRow(targetSheet)

    SendDashboardEmail targetSheet, rowCount
    Exit Sub

CleanFail:
    LogErr "EmailDashboard"
    MsgBox "An unexpected error occurred while handling the dashboard email button." & vbCrLf & vbCrLf & "See macro-debug.log for details.", vbExclamation, "Email Not Created"
End Sub
