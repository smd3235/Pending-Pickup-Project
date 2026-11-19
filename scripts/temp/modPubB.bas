Attribute VB_Name = "modPubB"
Option Explicit

Public Function Probe1() As Boolean
    gTestFlag = True
    gTestName = "x"
    Set gTestList = New Collection
    Probe1 = gTestFlag
End Function
