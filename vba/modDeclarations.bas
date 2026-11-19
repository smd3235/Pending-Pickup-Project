Attribute VB_Name = "modDeclarations"
Option Explicit

' ============================================================================
'  PENDING PICKUP REPORT BUILDER - v1.0
'  Module: modDeclarations
'  Holds all Public constants and shared module-level state used across
'  every other module in this project.
' ============================================================================

' ----------------------------------------------------------------------------
' Sheet / header constants
' ----------------------------------------------------------------------------
Public Const RAW_SHEET As String = "Raw Data"
Public Const MANAGER_HEADER As String = "Assigned Sales Manager"
Public Const STATUS_HEADER As String = "Treatment Status"
Public Const REP_HEADER As String = "Sales Representative"
Public Const NAME_HEADER As String = "Name"
Public Const KEY_HEADER As String = "EMPI"
' Alternative header names for EMPI / key column so older reports with different
' column names still match during prior-report comparison.
Public Const KEY_HEADER_ALIAS_1 As String = "Empl ID"
Public Const KEY_HEADER_ALIAS_2 As String = "Patient ID"
Public Const KEY_HEADER_ALIAS_3 As String = "Emp ID"
Public Const KEY_HEADER_ALIAS_4 As String = "Employee ID"
Public Const PICKUP_DATE_HEADER As String = "AUDIT: TP End Date"
Public Const START_DATE_HEADER As String = "Start Date"
Public Const PENDING_TEXT As String = "Pending Pickup"
Public Const MACRO_VERSION As String = "1.0"
Public Const MANAGER_HEADER_ROW As Long = 16

' ----------------------------------------------------------------------------
' File / log paths (derived at runtime from workbook location for portability)
' ----------------------------------------------------------------------------
Public Const AUDIT_SHEET_NAME As String = "Audit Log"
Public Const LOG_FILENAME As String = "macro-debug.log"
Public Const COMPLETION_MARKER_FILENAME As String = "pending-pickup-complete.txt"
Public Const AUTO_SOURCE_FILENAME As String = "auto-source.txt"
Public Const AUTO_SKIP_PRIOR_FILENAME As String = "auto-skip-prior.txt"

' Derive a writable log directory: workbook-relative first, then user-local fallback.
Public Function GetLogDirectory() As String
    On Error Resume Next
    Dim dirPath As String
    dirPath = Environ$("TEMP") & "\PendingPickupReport\log"
    If Len(Dir$(dirPath, vbDirectory)) = 0 Then MkDir dirPath
    GetLogDirectory = dirPath
End Function

Public Function GetLogPath() As String
    GetLogPath = GetLogDirectory() & "\" & LOG_FILENAME
End Function

Public Function GetCompletionMarkerPath() As String
    GetCompletionMarkerPath = GetLogDirectory() & "\" & COMPLETION_MARKER_FILENAME
End Function

Public Function GetAutoSourcePathFile() As String
    GetAutoSourcePathFile = GetLogDirectory() & "\" & AUTO_SOURCE_FILENAME
End Function

Public Function GetAutoSkipPriorFile() As String
    GetAutoSkipPriorFile = GetLogDirectory() & "\" & AUTO_SKIP_PRIOR_FILENAME
End Function

' ----------------------------------------------------------------------------
' Email constants
' ----------------------------------------------------------------------------
Public Const EMAIL_SUBJECT_MANAGER As String = "Pending Pickup Report - Action Required"
Public Const EMAIL_SUBJECT_DIRECTOR As String = "Pending Pickup Report - Monthly Dashboard"
Public Const EMAIL_REG_KEY As String = "PendingPickupReport\Email"

' Email mode / behavior toggles (read from registry under EMAIL_REG_KEY)
Public Const EMAIL_REG_AUTO_SEND As String = "AutoSend"        ' "1" = auto-send, "" or missing = draft only
Public Const EMAIL_REG_ATTACH_REPORT As String = "AttachReport" ' "1" = attach the generated report .xlsx
Public Const EMAIL_MODE_DRAFT As String = "DRAFT"
Public Const EMAIL_MODE_SEND As String = "SEND"

' {{REPORT}} marker: body text is split at this marker. Text above = intro,
' text below = closing. The copied Excel report range is pasted at the marker
' position. If the marker is absent, the full body is used as intro, the report
' is pasted after it, and a default closing is appended.
Public Const EMAIL_MARKER_REPORT As String = "{{REPORT}}"

' ----------------------------------------------------------------------------
' Manager email template/config sheet (hidden tab in template)
' ----------------------------------------------------------------------------
Public Const EMAIL_CONFIG_SHEET As String = "Manager Email Template"
Public Const EMAIL_CONFIG_TO_CELL As String = "D5"
Public Const EMAIL_CONFIG_CC_CELL As String = "D7"
Public Const EMAIL_CONFIG_SUBJECT_CELL As String = "D9"
Public Const EMAIL_CONFIG_BODY_CELL As String = "D11"
' Legacy START HERE cells (fallback)
Public Const LEGACY_EMAIL_CONFIG_SHEET As String = "START HERE"
Public Const LEGACY_EMAIL_CONFIG_CC_CELL As String = "D16"
Public Const LEGACY_EMAIL_CONFIG_SUBJECT_CELL As String = "D17"
Public Const LEGACY_EMAIL_CONFIG_BODY_CELL As String = "D18"

' ----------------------------------------------------------------------------
' Director email template/config sheet (hidden tab in template)
' ----------------------------------------------------------------------------
Public Const DIRECTOR_EMAIL_SHEET As String = "Director Email Template"
Public Const DIRECTOR_EMAIL_TO_CELL As String = "D5"
Public Const DIRECTOR_EMAIL_CC_CELL As String = "D7"
Public Const DIRECTOR_EMAIL_SUBJECT_CELL As String = "D9"
Public Const DIRECTOR_EMAIL_BODY_CELL As String = "D11"

' ----------------------------------------------------------------------------
' Shared runtime state (populated during each report run)
' ----------------------------------------------------------------------------
Public mAuditStartTime As Date
Public mAuditSourcePath As String
Public mAuditOutputPath As String
Public mAuditTotalRecords As Long
Public mAuditPendingRecords As Long
Public mAuditNewRecords As Long
Public mAuditClearedRecords As Long
Public mAuditStillRecords As Long
Public mAuditMissingDates As Long
Public mAuditDuplicateKeys As Long
Public mAuditMissingManagers As Long
Public mAuditMissingStatus As Long
Public mAuditFutureDates As Long
Public mAuditVeryOldDates As Long
Public mAuditErrorCount As Long
Public mAuditIssues As Collection
Public mRunLogPath As String
Public mInReportRun As Boolean
Public mRepSeenDict As Object

' ----------------------------------------------------------------------------
' Timing state
' ----------------------------------------------------------------------------
Public mReportStartTime As Date
