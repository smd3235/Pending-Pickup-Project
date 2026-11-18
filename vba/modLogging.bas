Attribute VB_Name = "modLogging"
Option Explicit

' ============================================================================
'  Module: modLogging
'  Centralised debug / audit logging. Every run writes to LOG_PATH and,
'  during a report run, also to the per-run log in mRunLogPath.
' ============================================================================

Public Sub LogMsg(ByVal msg As String)
    Dim f As Integer
    Dim logLine As String
    On Error Resume Next
    Err.Clear
    f = FreeFile
    If Err.Number <> 0 Then Exit Sub
    logLine = Format(Now, "yyyy-mm-dd hh:n:ss") & " | " & msg
    If Len(Dir$(Left$(LOG_PATH, InStrRev(LOG_PATH, "\")), vbDirectory)) = 0 Then MkDir Left$(LOG_PATH, InStrRev(LOG_PATH, "\"))
    Err.Clear
    Open LOG_PATH For Append As #f
    If Err.Number = 0 Then
        Print #f, logLine
        Close #f
    End If

    ' Also write to the per-run log when a report run is active so traces are run-scoped.
    If mInReportRun And Len(mRunLogPath) > 0 Then
        Err.Clear
        f = FreeFile
        If Err.Number = 0 Then
            Open mRunLogPath For Append As #f
            If Err.Number = 0 Then
                Print #f, logLine
                Close #f
            End If
        End If
    End If
End Sub

Public Sub LogErr(ByVal procName As String)
    Dim f As Integer
    Dim logLine As String
    On Error Resume Next
    Err.Clear
    f = FreeFile
    If Err.Number <> 0 Then Exit Sub
    logLine = Format(Now, "yyyy-mm-dd hh:n:ss") & " | " & procName & " | Err=" & Err.Number & " | " & Err.Description & " | Erl=" & Erl
    If Len(Dir$(Left$(LOG_PATH, InStrRev(LOG_PATH, "\")), vbDirectory)) = 0 Then MkDir Left$(LOG_PATH, InStrRev(LOG_PATH, "\"))
    Err.Clear
    Open LOG_PATH For Append As #f
    If Err.Number = 0 Then
        Print #f, logLine
        Close #f
    End If

    ' Also write to the per-run log when a report run is active.
    If mInReportRun And Len(mRunLogPath) > 0 Then
        Err.Clear
        f = FreeFile
        If Err.Number = 0 Then
            Open mRunLogPath For Append As #f
            If Err.Number = 0 Then
                Print #f, logLine
                Close #f
            End If
        End If
    End If
End Sub

Public Function BuildRunLogPath() As String
    Dim logDir As String
    On Error Resume Next
    logDir = Left$(LOG_PATH, InStrRev(LOG_PATH, "\"))
    BuildRunLogPath = logDir & "pending-pickup_" & Format(Now, "yyyy-mm-dd_HHMMSS") & ".log"
End Function

Public Sub WriteCompletionMarker(ByVal savePath As String)
    Dim f As Integer
    On Error Resume Next
    If Len(Dir$(Left$(COMPLETION_MARKER_PATH, InStrRev(COMPLETION_MARKER_PATH, "\")), vbDirectory)) = 0 Then MkDir Left$(COMPLETION_MARKER_PATH, InStrRev(COMPLETION_MARKER_PATH, "\"))
    f = FreeFile
    If Err.Number <> 0 Then Exit Sub
    Open COMPLETION_MARKER_PATH For Output As #f
    If Err.Number <> 0 Then Exit Sub
    Print #f, "completed=" & Format(Now, "yyyy-mm-dd hh:n:ss")
    Print #f, "saved=" & savePath
    Close #f
End Sub
