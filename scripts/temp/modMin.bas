Attribute VB_Name = "modMin"
Option Explicit

Public Sub LogMsg(ByVal msg As String)
    If mInReportRun And Len(mRunLogPath) > 0 Then
        Debug.Print mRunLogPath
    End If
End Sub
