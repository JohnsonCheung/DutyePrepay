Attribute VB_Name = "nStr_SCL"
Option Compare Database
Option Explicit

Function SclSy(SemiColonLin$, Optional NoTrim As Boolean) As String()
Dim A$(): A = Split(SemiColonLin, ";")
If Not NoTrim Then A = AyTrim(A)
SclSy = A
End Function
