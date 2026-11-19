Attribute VB_Name = "modFn"
Option Explicit

Public Function F() As String
    F = Environ$("TEMP")
End Function

Public mInReportRun As Boolean
