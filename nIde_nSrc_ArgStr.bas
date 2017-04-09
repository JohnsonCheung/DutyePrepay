Attribute VB_Name = "nIde_nSrc_ArgStr"
Option Compare Database
Option Explicit

Function ArgStrToNm$(ArgStr$)
Dim A$: A = ArgStr
Dim Dmy$
Dmy = ParseStr(A, "Optional ")
Dmy = ParseStr(A, "Paramarray ")
ArgStrToNm = ParseNm(A)
End Function
