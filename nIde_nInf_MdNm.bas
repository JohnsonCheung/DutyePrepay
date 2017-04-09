Attribute VB_Name = "nIde_nInf_MdNm"
Option Compare Database
Option Explicit

Function MdNmNz$(A$)
If A <> "" Then
    MdNmNz = A
Else
    MdNmNz = MdNm(MdCur)
End If
End Function
