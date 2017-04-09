Attribute VB_Name = "nDao_nDta_Tny"
Option Compare Database
Option Explicit

Function TnyDs(Tny$(), Optional A As database) As Ds
Dim J&
Dim O As Ds
For J = 0 To UB(Tny)
    O = DsAddDt(O, TblDt(Tny(J), , A))
Next
TnyDs = O
End Function
