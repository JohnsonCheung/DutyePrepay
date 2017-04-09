Attribute VB_Name = "nDao_nDta_Db"
Option Compare Database
Option Explicit

Function DbDs(Tny$(), Optional A As database) As Ds
Dim D As database: Set D = DbNz(A)
Dim J%, O As Ds
For J = 0 To UB(Tny)
    O = DsAddDt(O, TblDt(Tny(J), , D))
Next
DbDs = O
End Function

Sub DbDs__Tst()
'1 Declare
Dim Tny$()
Dim A As database
Dim Act As Ds
Dim Exp As Ds

'2 Assign
Tny = DbTny
Set A = CurrentDb

'3 Calling
Act = DbDs(Tny, A)

'4 Asst
Stop
Dim T$
    T = TmpFt
DsWrt Act, T
FtBrw T
End Sub
