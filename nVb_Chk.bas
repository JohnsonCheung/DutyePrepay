Attribute VB_Name = "nVb_Chk"
Option Compare Database
Option Explicit

Sub ChkBrw(Chk As Dt)
'[Chk] is a table with or without records.  If no record, it means nothing to check.
'If there is record, something needs to be checked.
'So ChkBrw is written as this.
If DtIsNoRec(Chk) Then Exit Sub
DtBrw Chk, "Please-check"
Err.Raise 1
End Sub
