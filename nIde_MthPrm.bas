Attribute VB_Name = "nIde_MthPrm"
Option Compare Database
Option Explicit

Function MthPrm(Mth As Mth, PrmAy() As Prm) As MthPrm
Dim O As New MthPrm
Set O.Mth = Mth
O.SetPrmAy PrmAy
Set MthPrm = O
End Function

Function MthPrmToStr$(A As MthPrm)
MthPrmToStr = MthToStr(A.Mth) & "(" & PrmAyToStr(A.PrmAy)
End Function
