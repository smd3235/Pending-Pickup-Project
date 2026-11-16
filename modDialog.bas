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
' Email config (read from START HERE cells on the template)
' ----------------------------------------------------------------------------

Public Function GetEmailCC() As String
    On Error Resume Next
    GetEmailCC = ThisWorkbook.Worksheets(EMAIL_CONFIG_SHEET).Range(EMAIL_CONFIG_CC_CELL).Value
    If Err.Number <> 0 Then GetEmailCC = ""
    Err.Clear
    LogMsg "Email config read: CC=" & GetEmailCC()
End Function

Public Function GetEmailSubject(ByVal defaultSubject As String) As String
    On Error Resume Next
    GetEmailSubject = ThisWorkbook.Worksheets(EMAIL_CONFIG_SHEET).Range(EMAIL_CONFIG_SUBJECT_CELL).Value
    If Err.Number <> 0 Or Len(GetEmailSubject) = 0 Then GetEmailSubject = defaultSubject
    Err.Clear
    LogMsg "Email config read: Subject=" & GetEmailSubject
End Function

Public Function GetEmailBody(ByVal defaultBody As String) As String
    On Error Resume Next
    GetEmailBody = ThisWorkbook.Worksheets(EMAIL_CONFIG_SHEET).Range(EMAIL_CONFIG_BODY_CELL).Value
    If Err.Number <> 0 Or Len(GetEmailBody) = 0 Then GetEmailBody = defaultBody
    Err.Clear
    LogMsg "Email config read: Body=" & Left$(GetEmailBody, 80)
End Function

Public Sub SetupEmailCC()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(EMAIL_CONFIG_SHEET)
    ws.Activate
    ws.Range(EMAIL_CONFIG_CC_CELL).Select
    LogMsg "Email setup: user opened CC configuration on START HERE"
    MsgBox "Edit the CC recipients in cell " & EMAIL_CONFIG_CC_CELL & " on the START HERE sheet." & vbCrLf & _
           "Enter comma-separated email addresses, then press Enter." & vbCrLf & vbCrLf & _
           "Subject and body can be edited in cells " & EMAIL_CONFIG_SUBJECT_CELL & " and " & EMAIL_CONFIG_BODY_CELL & ".", vbInformation
End Sub

' ----------------------------------------------------------------------------
' Email
' ----------------------------------------------------------------------------

' ----------------------------------------------------------------------------
' Helper: wait for Outlook Word editor to be available after opening inspector
' ----------------------------------------------------------------------------
Private Function GetWordEditor(ByVal mail As Object, Optional ByVal maxWaitSeconds As Single = 4) As Object
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
    Do While Timer - startTime < maxWaitSeconds
End Function

' ----------------------------------------------------------------------------
' Send email to a manager using copy/paste of Excel range into Outlook Word editor
' ----------------------------------------------------------------------------

Public Sub SendManagerEmail(ByVal ws As Worksheet, ByVal managerName As String, ByVal rowCount As Long)
    On Error GoTo CleanFail

    Dim snapRg As Range
    Set snapRg = ws.Range(ws.Cells(MANAGER_HEADER_ROW, 1), ws.Cells(LastUsedRow(ws), LastUsedColumn(ws)))

    LogMsg "Email: preparing Excel content for manager " & managerName & " (" & rowCount & " pending rows)"

    If ws Is Nothing Or snapRg Is Nothing Then
        LogMsg "Email FAILED: invalid range for manager " & managerName
        Exit Sub
    End If

    ' Copy the Excel range to the clipboard so formatting and spacing are preserved when pasted into Outlook.
    snapRg.Copy
    If Err.Number <> 0 Then
        LogMsg "Email FAILED: could not copy range for manager " & managerName & " (Err=" & Err.Number & ")"
        Err.Clear
        Exit Sub
    End If

    Dim outlook As Object, mail As Object
    Set outlook = CreateObject("Outlook.Application")
    If outlook Is Nothing Then
        LogMsg "Email FAILED: Outlook not available for manager " & managerName
        Exit Sub
    End If
    LogMsg "Email: Outlook application created for manager " & managerName
    Set mail = outlook.CreateItem(0)

    With mail
        .Subject = GetEmailSubject(EMAIL_SUBJECT_MANAGER & " - " & managerName & " (" & rowCount & " pending)")
        .CC = GetEmailCC()

        ' Display the inspector so we can access the Word editor and paste formatted Excel content.
        .Display
        DoEvents

        Dim wordDoc As Object
        Set wordDoc = GetWordEditor(mail, 4)
        If wordDoc Is Nothing Then
            LogMsg "Email FAILED: Word editor not available for manager " & managerName
            Exit Sub
        End If

        ' Insert intro text at the top.
        Dim bodyRange As Object
        Set bodyRange = wordDoc.Range(0, 0)
        bodyRange.Text = "Hi " & managerName & "," & vbCrLf & _
                         "Please review the Pending Pickup report below. You have " & rowCount & " records requiring pickup action." & vbCrLf & vbCrLf

        ' Move to the end and paste the copied Excel range (preserves formatting).
        Set bodyRange = wordDoc.Range(wordDoc.Characters.Count, wordDoc.Characters.Count)
        bodyRange.Paste

        ' Add closing text.
        Set bodyRange = wordDoc.Range(wordDoc.Characters.Count, wordDoc.Characters.Count)
        bodyRange.Text = vbCrLf & vbCrLf & "Please review and respond as needed." & vbCrLf & "Pending Pickup Report System"

        LogMsg "Email: inspector ready for manager " & managerName & " (CC=" & GetEmailCC() & ")"
    End With

    LogMsg "Email: inspector displayed for manager " & managerName
    Exit Sub

CleanFail:
    LogErr "SendManagerEmail"
End Sub

' ----------------------------------------------------------------------------
' Send dashboard email to Director using copy/paste of Excel range into Outlook Word editor
' ----------------------------------------------------------------------------

Public Sub SendDashboardEmail(ByVal ws As Worksheet, ByVal rowCount As Long)
    On Error GoTo CleanFail

    Dim snapRg As Range
    Set snapRg = ws.Range("A1:P25")

    LogMsg "Email: preparing Excel content for dashboard"

    If ws Is Nothing Or snapRg Is Nothing Then
        LogMsg "Email FAILED: invalid range for dashboard"
        Exit Sub
    End If

    ' Copy the Excel range to the clipboard so formatting and spacing are preserved when pasted into Outlook.
    snapRg.Copy
    If Err.Number <> 0 Then
        LogMsg "Email FAILED: could not copy range for dashboard (Err=" & Err.Number & ")"
        Err.Clear
        Exit Sub
    End If

    Dim outlook As Object, mail As Object
    Set outlook = CreateObject("Outlook.Application")
    If outlook Is Nothing Then
        LogMsg "Email FAILED: Outlook not available for dashboard email"
        Exit Sub
    End If
    LogMsg "Email: Outlook application created for dashboard email"
    Set mail = outlook.CreateItem(0)

    With mail
        .Subject = GetEmailSubject(EMAIL_SUBJECT_DIRECTOR)
        .CC = GetEmailCC()

        ' Display the inspector so we can access the Word editor and paste formatted Excel content.
        .Display
        DoEvents

        Dim wordDoc As Object
        Set wordDoc = GetWordEditor(mail, 4)
        If wordDoc Is Nothing Then
            LogMsg "Email FAILED: Word editor not available for dashboard email"
            Exit Sub
        End If

        ' Insert intro text at the top.
        Dim bodyRange As Object
        Set bodyRange = wordDoc.Range(0, 0)
        bodyRange.Text = "Hi Director," & vbCrLf & _
                         "Please find the Pending Pickup daily dashboard below." & vbCrLf & vbCrLf

        ' Move to the end and paste the copied Excel range (preserves formatting).
        Set bodyRange = wordDoc.Range(wordDoc.Characters.Count, wordDoc.Characters.Count)
        bodyRange.Paste

        ' Add closing text.
        Set bodyRange = wordDoc.Range(wordDoc.Characters.Count, wordDoc.Characters.Count)
        bodyRange.Text = vbCrLf & vbCrLf & "Pending Pickup Report System"

        LogMsg "Email: inspector ready for dashboard (CC=" & GetEmailCC() & ")"
    End With

    LogMsg "Email: inspector displayed for dashboard"
    Exit Sub

CleanFail:
    LogErr "SendDashboardEmail"
End Sub

' ----------------------------------------------------------------------------
' Button click handlers - use Application.Caller + AlternativeText to identify target sheet
' ----------------------------------------------------------------------------

Public Sub EmailSelectedManagerTab()
    On Error GoTo CleanFail
    LogMsg "Email: user clicked manager tab email button"

    ' Use Application.Caller to identify which button was clicked, then read its
    ' AlternativeText to get the target sheet name. This avoids relying on a
    ' stale global that gets overwritten during manager tab creation.
    Dim callerName As String
    callerName = Application.Caller
    If Len(callerName) = 0 Then
        LogMsg "Email FAILED: no caller identified for manager tab email button"
        Exit Sub
    End If

    Dim shp As Shape
    Set shp = Nothing
    On Error Resume Next
    Set shp = ThisWorkbook.Shapes(callerName)
    On Error GoTo CleanFail
    If shp Is Nothing Then
        ' Try on the active sheet
        Set shp = ActiveSheet.Shapes(callerName)
    End If
    If shp Is Nothing Then
        LogMsg "Email FAILED: could not find shape '" & callerName & "'"
        Exit Sub
    End If

    Dim targetSheetName As String
    targetSheetName = shp.AlternativeText
    If Len(targetSheetName) = 0 Then
        LogMsg "Email FAILED: shape '" & callerName & "' has no target sheet in AlternativeText"
        Exit Sub
    End If

    Dim targetSheet As Worksheet
    On Error Resume Next
    Set targetSheet = ThisWorkbook.Worksheets(targetSheetName)
    On Error GoTo CleanFail
    If targetSheet Is Nothing Then
        LogMsg "Email FAILED: worksheet '" & targetSheetName & "' not found"
        Exit Sub
    End If

    ' Determine manager name from shape text (e.g. "Email Victor Alvarez")
    Dim mgrName As String
    mgrName = Replace(shp.TextFrame2.TextRange.Text, "Email ", "")

    ' Count data rows on the manager sheet
    Dim rowCount As Long
    rowCount = LastUsedRow(targetSheet) - MANAGER_HEADER_ROW
    If rowCount < 0 Then rowCount = 0

    SendManagerEmail targetSheet, mgrName, rowCount
    Exit Sub

CleanFail:
    LogErr "EmailSelectedManagerTab"
End Sub

Public Sub EmailDashboard()
    On Error GoTo CleanFail
    LogMsg "Email: user clicked dashboard email button"

    Dim callerName As String
    callerName = Application.Caller
    Dim shp As Shape
    Set shp = Nothing
    On Error Resume Next
    Set shp = ThisWorkbook.Shapes(callerName)
    On Error GoTo CleanFail
    If shp Is Nothing Then
        Set shp = ActiveSheet.Shapes(callerName)
    End If

    Dim targetSheet As Worksheet
    If Not shp Is Nothing And Len(shp.AlternativeText) > 0 Then
        On Error Resume Next
        Set targetSheet = ThisWorkbook.Worksheets(shp.AlternativeText)
        On Error GoTo CleanFail
    End If
    If targetSheet Is Nothing Then
        ' Fallback: find the Dashboard sheet
        On Error Resume Next
        Set targetSheet = ThisWorkbook.Worksheets("Dashboard")
        On Error GoTo CleanFail
    End If
    If targetSheet Is Nothing Then
        LogMsg "Email FAILED: could not find Dashboard worksheet"
        Exit Sub
    End If

    Dim rowCount As Long
    rowCount = LastUsedRow(targetSheet)
    SendDashboardEmail targetSheet, rowCount
    Exit Sub

CleanFail:
    LogErr "EmailDashboard"
End Sub
