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
Public Const PICKUP_DATE_HEADER As String = "AUDIT: TP End Date"
Public Const START_DATE_HEADER As String = "Start Date"
Public Const PENDING_TEXT As String = "Pending Pickup"
Public Const MACRO_VERSION As String = "1.0"
Public Const MANAGER_HEADER_ROW As Long = 16

' ----------------------------------------------------------------------------
' File / log paths
' ----------------------------------------------------------------------------
Public Const LOG_PATH As String = "C:\Users\smd32\Downloads\Pending Pickup Project\log\macro-debug.log"
Public Const COMPLETION_MARKER_PATH As String = "C:\Users\smd32\Downloads\Pending Pickup Project\log\pending-pickup-complete.txt"
Public Const AUDIT_SHEET_NAME As String = "Audit Log"
Public Const AUTO_SOURCE_PATH_FILE As String = "C:\Users\smd32\Downloads\Pending Pickup Project\log\auto-source.txt"
Public Const AUTO_SKIP_PRIOR_FILE As String = "C:\Users\smd32\Downloads\Pending Pickup Project\log\auto-skip-prior.txt"

' ----------------------------------------------------------------------------
' Email constants
' ----------------------------------------------------------------------------
Public Const EMAIL_SUBJECT_MANAGER As String = "Pending Pickup Report - Action Required"
Public Const EMAIL_SUBJECT_DIRECTOR As String = "Pending Pickup Report - Daily Dashboard"
Public Const EMAIL_REG_KEY As String = "PendingPickupReport\Email"

' ----------------------------------------------------------------------------
' Email config cells on START HERE template sheet (manager emails)
' ----------------------------------------------------------------------------
Public Const EMAIL_CONFIG_SHEET As String = "START HERE"
Public Const EMAIL_CONFIG_CC_CELL As String = "D16"
Public Const EMAIL_CONFIG_SUBJECT_CELL As String = "D17"
Public Const EMAIL_CONFIG_BODY_CELL As String = "D18"

' ----------------------------------------------------------------------------
' Director email template/config sheet (hidden tab in template)
' ----------------------------------------------------------------------------
Public Const DIRECTOR_EMAIL_SHEET As String = "Director Email Template"
Public Const DIRECTOR_EMAIL_TO_CELL As String = "D5"
Public Const DIRECTOR_EMAIL_CC_CELL As String = "D6"
Public Const DIRECTOR_EMAIL_SUBJECT_CELL As String = "D7"
Public Const DIRECTOR_EMAIL_BODY_CELL As String = "D8"

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
