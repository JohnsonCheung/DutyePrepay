Attribute VB_Name = "nIde_nSrc_PrmStr"
Option Compare Database
Option Explicit

Function PrmStrToNy(PrmStr$) As String()
Dim A$(): A = AyTrim(Split(PrmStr, ","))
PrmStrToNy = AyMapIntoSy(A, "ArgStrToNm")
End Function

Sub PrmStrToNy__Tst()
AyAsstEq PrmStrToNy("PrmStr$()"), ApSy("PrmStr")
End Sub
