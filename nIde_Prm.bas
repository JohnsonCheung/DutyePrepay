Attribute VB_Name = "nIde_Prm"
Option Compare Database
Option Explicit

Function PrmBrk(PrmStr$) As String()
Dim PrmNm$, PrmSfx$
    Dim J%, C$
    For J = 1 To Len(PrmStr)
        C = Mid(PrmStr, J, 1)
        If ChrIsNmChr(C) Then
            PrmNm = PrmNm & C
        Else
            PrmSfx = Mid(PrmStr, J)
            Exit For
        End If
    Next
PrmBrk = ApSy(PrmNm, PrmSfx)
End Function

Sub PrmBrk__Tst()
Dim PrmStr$
Dim Act$()
Dim Exp$()

PrmStr = "PrmStr$"
Act = PrmBrk(PrmStr)
Exp = ApSy("PrmStr", "$")

AyAsstEq Act, Exp
End Sub
