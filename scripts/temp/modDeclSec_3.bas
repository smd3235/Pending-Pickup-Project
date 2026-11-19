Attribute VB_Name = "modDeclSec"
Option Explicit

Public Function GetAutoSourcePathFile() As String
    GetAutoSourcePathFile = GetLogDirectory() & "\" & AUTO_SOURCE_FILENAME
End Function

Public mInReportRun As Boolean
