Attribute VB_Name = "nDao_nPrimeTbl_Tbl"
Option Compare Database
Option Explicit

Function TblIsPrime(T$, Optional A As database) As Boolean
Dim D As database: Set D = DbNz(A)
Dim I As DAO.Index: Set I = TblPriIdx(T, D)
If IsNothing(I) Then Exit Function
Dim F As IndexFields
Set F = I.Fields
If F.Count <> 1 Then Exit Function
Dim F2 As Field2
Set F2 = F(0)
If F2.Name <> T Then Exit Function ' Er "{Tbl} has Single-Field-Primary-Idx with {Idx-Nm} does not same as Tbl", T, F2.Name
TblIsPrime = True
End Function
