Attribute VB_Name = "nIde_nTst_Md"
Option Compare Database
Option Explicit

Function MdIsTstNm(Optional A As CodeModule) As Boolean
MdIsTstNm = NmIsTstNm(MdNm(A))
End Function

