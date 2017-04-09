Attribute VB_Name = "nVb_Dte"
Option Compare Database
Option Explicit

Function ChkMth$(M As Byte)
If M > 12 Or M < 1 Then ChkMth = FmtQQ("Given [M] must between 1 and 12.  [M]=[?]", M)
End Function

Function ChkYr$(Y As Byte)
If Y = 0 Then ChkYr = "Given [Y] must >0"
End Function

Function DteYMD(D As Date) As YMD
Dim O As YMD
With O
    .Y = CInt(Year(D) - 2000)
    .M = Month(D)
    .D = Day(D)
End With
DteYMD = O
End Function

Function VdtMth(M As Byte, Optional NoMsg As Boolean) As Boolean
Dim A$
    A = ChkMth(M)
    If A = "" Then Exit Function
VdtMth = True
If NoMsg Then Exit Function
MsgBox A, vbCritical
End Function

Function VdtYr(Y As Byte, Optional NoMsg As Boolean) As Boolean
Dim A$
    A = ChkYr(Y)
    If A = "" Then Exit Function
VdtYr = True
If NoMsg Then Exit Function
MsgBox A, vbCritical
End Function
