Attribute VB_Name = "modDeclTest"
Option Explicit

Public mInReportRun As Boolean
Private moIssues As Collection
Private moRepSeen As Object

Public Property Get AuditIssues() As Collection
    Set AuditIssues = moIssues
End Property

Public Property Set AuditIssues(ByVal v As Collection)
    Set moIssues = v
End Property

Public Property Get RepSeenDict() As Object
    Set RepSeenDict = moRepSeen
End Property

Public Property Set RepSeenDict(ByVal v As Object)
    Set moRepSeen = v
End Property
