Attribute VB_Name = "modDeclSec"
Option Explicit

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

Public mInReportRun As Boolean
