Attribute VB_Name = "modDeclSec"
Option Explicit

Public Function GetLogPath() As String
    GetLogPath = GetLogDirectory() & "\" & LOG_FILENAME
End Function

Public mInReportRun As Boolean
