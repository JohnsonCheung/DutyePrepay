Attribute VB_Name = "nDta_nScl_Ds"
Option Compare Database
Option Explicit

Function DsRead(Ft) As Ds
Dim Gp(): Gp = AyGpByPfx(FtLy(Ft), "Tbl;")
Dim O As Ds, J%
For J = 0 To UB(Gp)
    O = DsAddDt(O, DtNewScLy(Gp(J)))
Next
DsRead = O
End Function

Sub DsWrt(A As Ds, Ft$)
If DsIsEmpty(A) Then StrWrt "", Ft: Exit Sub
Dim J%
Dim F%: F = FtOpnOup(Ft)
For J = 0 To DsUTbl(A)
    AyWrtFno DtScLy(DsDt(A, J)), F
Next
Close #F
End Sub
