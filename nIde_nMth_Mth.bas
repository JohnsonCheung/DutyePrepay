Attribute VB_Name = "nIde_nMth_Mth"
Option Compare Database
Option Explicit

Function Mth(MdNm$, MthNm$) As Mth
Dim O As New Mth
O.MdNm = MdNm
O.MthNm = MthNm
Set Mth = O
End Function

Function MthToStr$(A As Mth)
MthToStr = A.MdNm & "." & A.MthNm
End Function
