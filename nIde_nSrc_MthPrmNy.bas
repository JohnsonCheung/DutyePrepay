Attribute VB_Name = "nIde_nSrc_MthPrmNy"
Option Compare Database
Option Explicit

Function MthPrmNy(MthNm$, Optional A As CodeModule) As String()
MthPrmNy = PrmStrToNy(MthBrkNew(MthLin(MthNm, A)).PrmStr)
End Function

