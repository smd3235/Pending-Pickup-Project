Attribute VB_Name = "modDeclSec"
Option Explicit

Public Function GetCompletionMarkerPath() As String
    GetCompletionMarkerPath = GetLogDirectory() & "\" & COMPLETION_MARKER_FILENAME
End Function

Public mInReportRun As Boolean
