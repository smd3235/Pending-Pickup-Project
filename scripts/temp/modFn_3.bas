Attribute VB_Name = "modFn"
Option Explicit

Public Function F() As String
    If Len(Dir$("C:\", vbDirectory)) = 0 Then F = ""
End Function

Public mInReportRun As Boolean
