Attribute VB_Name = "modDeclTest"
Option Explicit

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
Public mReportStartTime As Date
