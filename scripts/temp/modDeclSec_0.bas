Attribute VB_Name = "modDeclSec"
Option Explicit

Public Function GetLogDirectory() As String
    On Error Resume Next
    Dim dirPath As String
    dirPath = Environ$("TEMP") & "\PendingPickupReport\log"
    If Len(Dir$(dirPath, vbDirectory)) = 0 Then MkDir dirPath
    GetLogDirectory = dirPath
End Function

Public mInReportRun As Boolean
