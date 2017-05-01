Attribute VB_Name = "nStr_Rmv"
Option Compare Database
Option Explicit

Function Rmv_SqBkt$(pS$)
If Left(pS, 1) = "[" And Right(pS, 1) = "]" Then Rmv_SqBkt = Mid(pS, 2, Len(pS) - 2): Exit Function
Rmv_SqBkt = pS
End Function

Function RmvBlankChr$(S)
RmvBlankChr = RmvLF(RmvCR(RmvTab(RmvSpc(S))))
End Function

Function RmvCR$(S)
RmvCR = Replace(S, vbCr, "")
End Function

Function RmvDoubleBlackSlash$(pStr$)
Dim mA$, mP%
mA = pStr
mP = InStr(pStr, "\\")
While mP > 0
    mA = Replace(mA, "\\", "\")
    mP = InStr(mA, "\\")
Wend
RmvDoubleBlackSlash = mA
End Function

Function RmvFstChr$(S)
RmvFstChr = Mid(S, 2)
End Function

Function RmvLasChr$(S)
RmvLasChr = Left(S, Len(S) - 1)
End Function

Function RmvLF$(S$)
RmvLF = Replace(S, vbLf, "")
End Function

Function RmvPfx$(S, Pfx)
If IsPfx(S, Pfx) Then
    RmvPfx = Mid(S, Len(Pfx) + 1)
Else
    RmvPfx = S
End If
End Function

Sub RmvPfx__Tst()
Debug.Assert RmvPfx("Tst_aaa", "Tst_") = "aaa"
Debug.Assert RmvPfx("Tst__aaa", "Tst__") = "aaa"
End Sub

Function RmvPfxAll$(S, Pfx)
Dim O$: O = S
Dim P%: P = Len(Pfx) + 1
While IsPfx(O, Pfx)
    O = Mid(O, P)
Wend
RmvPfxAll = O
End Function

Function RmvSfx$(S, Sfx)
If IsSfx(S, Sfx) Then
    RmvSfx = Left(S, Len(S) - Len(Sfx))
Else
    RmvSfx = S
End If
End Function

Sub RmvSfx__Tst()
Debug.Assert RmvSfx("aaa_Tst", "_Tst") = "aaa"
Debug.Assert RmvSfx("aaa__Tst", "__Tst") = "aaa"
End Sub

Function RmvSpc$(S)
RmvSpc = Replace(S, " ", "")
End Function

Function RmvTab$(S)
RmvTab = Replace(S, vbTab, "")
End Function
