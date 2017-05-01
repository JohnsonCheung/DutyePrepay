Attribute VB_Name = "nStr_Brk"
Option Compare Database
Option Explicit

Function Brk(S, BrkChr$, Optional NoTrim As Boolean) As S1S2
'Aim: Brk {S} into {S1S2}  Format of pS: <oS1><pBrkChr><oS2>) with both <oS1> & <oS2> & <pBrkChr> must exist
If BrkChr = "" Then Er "BrkChr must be given"
Dim At&: At = InStr(S, BrkChr): If At = 0 Then Er "{S} must contain {BrkChr}", S, BrkChr
Dim O As S1S2
    O = BrkAt(S, At, Len(BrkChr))
    If Not NoTrim Then O = S1S2Trim(O)
Brk = O
End Function

Function Brk_Brk_Cmd(oBrk$, oSplit$, OInto$, oTo$, oKeep$, oSetSno$, oBeg%, oStp%, pBrkCmd$) As Boolean
'Aim: Break {pBrkCmd} into: Brk Split [To] [Into] [Keep] [SetSno] [Beg] [Stp]
'     Assume no space with the elements
Const cSub$ = "Brk_Brk_Cmd"
Dim mBrkCmd$: mBrkCmd = Replace(Replace(Replace(pBrkCmd, vbLf, " "), vbCr, " "), "  ", " ")
Dim mA$(): mA = Split(mBrkCmd)
Dim J%
oBrk = "": oSplit = "": oTo = "": OInto = "": oKeep = "": oSetSno = "": oStp = 0
For J = 0 To Sz(mA) - 1 Step 2
    Select Case mA(J)
    Case "Brk":     If oBrk <> "" Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oBrk = mA(J + 1)
    Case "Split":   If oSplit <> "" Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oSplit = mA(J + 1)
    Case "To":      If oTo <> "" Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oTo = mA(J + 1)
    Case "Into":    If OInto <> "" Then ss.A 1, mA(J) & " is more than one": GoTo E
                       OInto = mA(J + 1)
    Case "Keep":    If oKeep <> "" Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oKeep = mA(J + 1)
    Case "SetSno":  If oSetSno <> "" Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oSetSno = mA(J + 1)
    Case "Beg":     If oBeg <> 0 Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oBeg = Val(mA(J + 1))
    Case "Stp":     If oStp <> 0 Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oStp = Val(mA(J + 1))
    Case Else
                    ss.A 1, "Element(" & mA(J) & ") is not expected.", , "Expected Values", "Brk Split Into To Keep SetSno Stp": GoTo E
    End Select
Next
Exit Function
E:
End Function

Function Brk_Cmb_Cmd(oCmb$, oJoin$, OInto$, oTo$, oKeep$, oOrd$, oStp%, pCmbCmd$) As Boolean
'Aim: Break {pJnCmd} into: Cmb Jn [To] [Into] [Keep] [Ord] [Stp]
'     Assume no space with the elements
Const cSub$ = "Brk_Cmb_Cmd"
Dim mCmbCmd$: mCmbCmd = Replace(Replace(Replace(pCmbCmd, vbLf, " "), vbCr, " "), "  ", " ")
Dim mA$(): mA = Split(mCmbCmd)
Dim J%
oCmb = "": oJoin = "": oTo = "": OInto = "": oKeep = "": oOrd = "": oStp = 0
For J = 0 To Sz(mA) - 1 Step 2
    Select Case mA(J)
    Case "Cmb":     If oCmb <> "" Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oCmb = mA(J + 1)
    Case "Join":    If oJoin <> "" Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oJoin = mA(J + 1)
    Case "To":      If oTo <> "" Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oTo = mA(J + 1)
    Case "Into":    If OInto <> "" Then ss.A 1, mA(J) & " is more than one": GoTo E
                       OInto = mA(J + 1)
    Case "Keep":    If oKeep <> "" Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oKeep = mA(J + 1)
    Case "Ord":     If oOrd <> "" Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oOrd = mA(J + 1)
    Case "Stp":     If oStp <> 0 Then ss.A 1, mA(J) & " is more than one": GoTo E
                       oStp = Val(mA(J + 1))
    Case Else
                    ss.A 1, "Element(" & mA(J) & ") is not expected is more than one": GoTo E
    End Select
Next
Exit Function
E:
End Function

Function Brk1(S, BrkChr$, Optional NoTrim As Boolean) As S1S2
If BrkChr = "" Then Er "BrkChr must be given"
Dim At&: At = InStr(S, BrkChr)
Brk1 = Brk1At(S, At, Len(BrkChr), NoTrim)
End Function

Function Brk1At(S, At&, L%, Optional NoTrim As Boolean) As S1S2
Dim O As S1S2
If At = 0 Then
    O.S1 = S
Else
    O = BrkAt(S, At, L)
End If
If Not NoTrim Then O = S1S2Trim(O)
Brk1At = O
End Function

Function Brk1FmEnd(S, BrkChr$, Optional NoTrim As Boolean) As S1S2
If BrkChr = "" Then Er "BrkChr must be given"
Dim At&: At = InStrRev(S, BrkChr)
Brk1FmEnd = Brk1At(S, At, Len(BrkChr), NoTrim)
End Function

Function Brk2(S, BrkChr$, Optional NoTrim As Boolean) As S1S2
If BrkChr = "" Then Er "BrkChr must be given"
Dim At&: At = InStr(S, BrkChr)
Brk2 = Brk2At(S, At, Len(BrkChr), NoTrim)
End Function

Function Brk2At(S, At&, L%, Optional NoTrim As Boolean) As S1S2
Dim O As S1S2
If At = 0 Then
    O.S2 = S
Else
    O = BrkAt(S, At, L)
End If
If Not NoTrim Then O = S1S2Trim(O)
Brk2At = O
End Function

Function Brk2FmEnd(S, BrkChr$, Optional NoTrim As Boolean) As S1S2
If BrkChr = "" Then Er "BrkChr must be given"
Dim At&: At = InStrRev(S, BrkChr)
Brk2FmEnd = Brk2At(S, At, Len(BrkChr), NoTrim)
End Function

Function BrkAt(S, At&, L%) As S1S2
If At = 0 Then Er "{At} cannot be 0", At
BrkAt.S1 = Left(S, At - 1)
BrkAt.S2 = Mid(S, At + L)
End Function

Function BrkCamelCas(CamelCasStr$) As String()
Dim J%
Dim O$(), S$, M$, A$
S = CamelCasStr
M = ParseFstChr(S)
While Len(S) > 0
    A = ParseFstChr(S)
    If ChrIsUcas(A) Then
        Push O, M
        M = A
    Else
        M = M & A
    End If
Wend
Push O, M
BrkCamelCas = O
End Function

Sub BrkCamelCas__Tst()
Dim A$
Dim Act$(), Exp$()
A = "CanYouComeHere"
Act = BrkCamelCas(A)
Exp = ApSy("Can", "You", "Come", "Here")
AyAsstEqExa Act, Exp

A = "A1A2A3A4A5A6A7"
Act = BrkCamelCas(A)
Exp = ApSy("A1", "A2", "A3", "A4", "A5", "A6", "A7")
AyAsstEqExa Act, Exp
End Sub

Function BrkFmEnd(S, BrkChr$, Optional NoTrim As Boolean) As S1S2
'Aim: Brk {S} into {S1S2}  Format of pS: <oS1><pBrkChr><oS2>) with both <oS1> & <oS2> & <pBrkChr> must exist
If BrkChr = "" Then Er "BrkChr must be given"
Dim At&: At = InStrRev(S, BrkChr): If At = 0 Then Er "{S} must contain {BrkChr}", S, BrkChr
Dim O As S1S2
    O = BrkAt(S, At, Len(BrkChr))
    If Not NoTrim Then O = S1S2Trim(O)
BrkFmEnd = O
End Function

Function BrkKWAy(S, KWAy$()) As String()
Dim O$, J%
O = RplCrLf(S)
For J = 0 To UB(KWAy)
    O = Replace(O, KWAy(J), vbCrLf & KWAy(J))
Next
BrkKWAy = Split(O, vbCrLf)
End Function

Function BrkMacroStr(MacroStr$) As String()
Dim A$(), P%
A = Split(MacroStr, "{")
Dim J%, O$()
For J = 1 To UB(A)
    PushNoBlank O, TakBef(A(J), "}")
Next
BrkMacroStr = O
End Function

Sub BrkMacroStr__Tst()
Dim A$: A = "sdf{abc} {def} dklsfj{xyz}"
Dim Act$(): Act = BrkMacroStr(A)
Dim Exp$(): Exp = ApSy("abc", "def", "xyz")
AyAsstEq Act, Exp
End Sub

Function BrkMacroStr1(MacroStr, Optional ExclBkt As Boolean) As String()
Dim A$(): A = Split(MacroStr, "{")
Dim O$(), J&, P&
For J = 1 To UB(A)
    P = InStr(A(J), "}")
    If P > 1 Then
        PushNoDup O, Left(A(J), P - 1)
    End If
Next
If Not ExclBkt Then O = AyQuote(O, "{}")
BrkMacroStr1 = O
End Function

Sub BrkMacroStr1__Tst()
AyAsstEq BrkMacroStr1("{a} b {d} {cccc}dd", ExclBkt:=True), ApSy("a", "d", "cccc")
AyAsstEq BrkMacroStr1("{a} b {d} {cccc}dd"), ApSy("{a}", "{d}", "{cccc}")
End Sub

Function BrkTo3Seg(oS1, oS2, oS3, pS$, Optional pBrkChr$ = ":", Optional pNoTrim As Boolean) As Boolean
Const cSub$ = "BrkTo3Seg"
Dim A$
If Brk_Str_0Or2(oS1, A, pS, pBrkChr, pNoTrim) Then ss.A 1: GoTo E
BrkTo3Seg = Brk_Str_0Or2(oS2, oS3, A, pBrkChr, pNoTrim)
Exit Function
R: ss.R
E:
End Function

Function BrkTo4Seg(oS1, oS2, oS3, oS4, pS$, Optional pBrkChr$ = ":", Optional pNoTrim As Boolean) As Boolean
Const cSub$ = "Brk_Str2Seg4"
Dim A$
If BrkTo3Seg(oS1, oS2, A, pS, pBrkChr, pNoTrim) Then ss.A 1: GoTo E
BrkTo4Seg = Brk_Str_0Or2(oS3, oS4, A, pBrkChr, pNoTrim)
Exit Function
R: ss.R
E:
End Function

Function BrkTo5Seg(oS1, oS2, oS3, oS4, oS5, pS$, Optional pBrkChr$ = ":", Optional pNoTrim As Boolean) As Boolean
Const cSub$ = "BrkTo5Seg"
Dim A$
If BrkTo4Seg(oS1, oS2, oS3, A, pS, pBrkChr, pNoTrim) Then ss.A 1: GoTo E
BrkTo5Seg = Brk_Str_0Or2(oS4, oS5, A, pBrkChr, pNoTrim)
Exit Function
R: ss.R
E:
End Function
