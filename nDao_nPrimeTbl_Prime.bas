Attribute VB_Name = "nDao_nPrimeTbl_Prime"
Option Compare Database
Option Explicit

Function PrimeTny(Optional A As database) As String()
Dim D As database: Set D = DbNz(A)
Dim Tny$(): Tny = DbTny(D)
PrimeTny = AySel(Tny, "TblIsPrime", D)
End Function
