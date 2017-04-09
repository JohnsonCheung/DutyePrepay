Attribute VB_Name = "nStr_Str"
Option Compare Database
Option Explicit

Function AddItmAft$(S, Itm$)
If S = "" Then Exit Function
AddItmAft = S & Itm
End Function

Function AddItmBef$(S, Itm$)
If S = "" Then Exit Function
AddItmBef = Itm & S
End Function

Function AddSpcAft$(S)
AddSpcAft = AddItmAft(S, " ")
End Function

Function AddSpcBef$(S)
AddSpcBef = AddItmBef(S, " ")
End Function

Function FstAsc%(S)
FstAsc = Asc(Left(S, 1))
End Function

Function FstChr$(S)
FstChr = Left(S, 1)
End Function

Function IsPfx(S, Pfx) As Boolean
IsPfx = Left(S, Len(Pfx)) = Pfx
End Function

Function IsPfxAp(S, ParamArray PfxAp()) As Boolean
Dim Av(): Av = PfxAp
Dim I
For Each I In Av
    If IsPfx(S, I) Then IsPfxAp = True: Exit Function
Next
End Function

Function IsSfx(S, Sfx) As Boolean
IsSfx = Right(S, Len(Sfx)) = Sfx
End Function

Function LasChr$(S)
LasChr = Right(S, 1)
End Function

Function RTrimSemiQ$(S)
Dim A$: A = Trim(S)
While Right(A, 1) = ";"
    A = RmvLasChr(A)
Wend
RTrimSemiQ = A
End Function

Function SplitVBar(VBarStr) As String()
SplitVBar = Split(VBarStr, "|")
End Function

Sub StrAsstEq(S1, S2)
ErAsst StrChkEq(S1, S2)
End Sub

Sub StrAsstEq__Tst()
StrAsstEq "slkfjsdfslkfslkfjsdfslkfslkfjsdfslkfslkfjsdfslkfslkfjsdfslkfslkfjsdfslkfslkfjsdfslkfjsdfslkfjsdfslkfjsdfslkfjsdf lsdkf sd", "slkfjsdfslkfslkfjsdfslkfslkfjsdfslkfslkfjsdfslkfslkfjsdfslkfslkfjsdfslkfslkfjsdfslkfjsdfslkfjsdfslkfjsdflskdfjdlf"
End Sub

Sub StrBrw(S, Optional Pfx$)
Dim T$: T = TmpFt(Pfx, "StrBrw")
StrWrt S, T
FtBrw T, True
End Sub

Function StrChkEq(S1, S2) As Dt
If S1 = S2 Then Exit Function
'==================
Dim L1&, L2&
L1 = Len(S1)
L2 = Len(S2)
'==================
Dim P&, A$, A1$, A2$
P = StrDifPos(S1, S2)
A = IIf(P = 1, "", "..")
A1 = IIf(P + 40 < L1, "", "..")
A2 = IIf(P + 40 < L2, "", "..")
'==================
Dim D$, D1$, D2$
If L1 <= 60 And L2 <= 60 Then
    D = "{Pos}: " & Space(P) & "v"
    D1 = S1
    D2 = S2
Else
    D1 = A & Mid(S1, P, 40) & A1
    D2 = A & Mid(S2, P, 40) & A2
End If
D1 = "{Len}: " & Quote(D1, "[]")
D2 = "{Len}: " & Quote(D2, "[]")
'==================
Dim O As Dt
O = ErNew("Two strings are different:")
O = ErApd(O, D, P)
O = ErApd(O, D1, L1)
O = ErApd(O, D2, L2)
StrChkEq = O
End Function

Sub StrChkEq__Tst()
Dim Er As Dt: Er = StrChkEq("1;2;3", "1;2;4")
DtBrw Er
End Sub

Function StrDifPos&(S1, S2)
If S1 = S2 Then Exit Function
Dim N&, J&
N = Min(Len(S1), Len(S2))
For J = 1 To N
    If Mid(S1, J, 1) <> Mid(S2, J, 1) Then StrDifPos = J: Exit Function
Next
StrDifPos = N
End Function

Function StrDup$(S, N%)
StrDup = String(N, S)
End Function

Function StrExpand$(S, SyOpt, Optional Sep$ = CtComma, Optional MacroStr$ = "{?}")
'Aim: Format a string with {?} by repeatly join it after substitue {?} by pAn$(0 to N)
StrExpand = Join(StrExpandToSy(S, SyOpt, MacroStr), Sep)
End Function

Sub StrExpand__Tst()
Dim Sy$()
Dim J%
For J = 0 To 3
    Push Sy, "[" & J & "]"
Next
Debug.Print StrExpand("xxxx{?}yyyyy", Sy, " and ")
Debug.Print StrExpand("Tbl{?}", "x xx xxx")
Debug.Print StrExpand("Tbl{?}", "x xx xxx", vbCrLf)

Dim mLines$:
mLines = RplVBar("Line1:lksdjflskdf sdklfj|" & _
"Line2:klsdjf{?}klsdjf}|" & _
"Line3:ksldjfslkdf")
Sy = Split("<Itm1>,<Itm2>,<Itm3>,<Itm4>", ",")
Debug.Print StrExpand(mLines, Sy, vbCrLf)
End Sub

Function StrExpandSeq$(pBeg As Byte, pN As Byte, Optional pFmtStr$ = "{N}", Optional pSepChr$ = CtComma, Optional pMacroStr$ = "{N}")
'Aim: Build a string to repeating {pFmtStr} {pN} times from {pBeg} with separated by {pSepChr}.  {pFmtStr} has {N} as the Idx.
Const cSub$ = "StrExpandSeq"
Dim mA$, J As Byte
For J = pBeg To pBeg + pN - 1
    mA = Add_Str(mA, Replace(pFmtStr, pMacroStr, J), pSepChr)
Next
StrExpandSeq = mA
End Function

Sub StrExpandSeq__Tst()
Dim mExpr$
mExpr = "StrExpandSeq(0, 10, ""a{N} as xx{N}"")"
Debug.Print "================="
Debug.Print mExpr
Debug.Print Eval(mExpr) ' StrExpandSeq(1, 10, "a{N} as xx{N}")
Debug.Print
mExpr = "StrExpandSeq(1, 10, ""a{N} as xx{N}"")"
Debug.Print mExpr
Debug.Print Eval(mExpr) ' StrExpandSeq(1, 10, "a{N} as xx{N}")
End Sub

Function StrExpandToSy(S, SyOpt, Optional MacroStr$ = "{?}") As String()
'Aim: Format a string with {?} by repeatly join it after substitue {?} by pAn$(0 to N)
Dim OSy$(): OSy = OptSy(SyOpt)
If InStr(S, MacroStr) = 0 Then
    StrExpandToSy = ApSy(S)
    Exit Function
End If
Dim O$()
    Dim U&
    U = UB(OSy)
    ReDim O(U)
    Dim J&
    For J = 0 To U
        O(J) = Replace(S, MacroStr, OSy(J))
    Next
StrExpandToSy = O
End Function

Function StrHas(S, SubStr) As Boolean
StrHas = InStr(S, SubStr) > 0
End Function

Function StrIsBlank(S) As Boolean
StrIsBlank = RmvBlankChr(S) = ""
End Function

Function StrIsIn(SubStr, S) As Boolean
StrIsIn = InStr(S, SubStr)
End Function

Function StrIsLik(S, Lik$) As Boolean
StrIsLik = S Like Lik
End Function

Function StrIsLikAy(S, LikAy) As Boolean
Dim J%
For J = 0 To UB(LikAy)
    If S Like LikAy(J) Then StrIsLikAy = True: Exit Function
Next
End Function

Function StrIsLikAy__Tst()
Debug.Assert StrIsLikAy("aa", Array("bb*", "cc*")) = False
Debug.Assert StrIsLikAy("aa", Array("bb*", "aa*")) = True
End Function

Function StrIsMacro(pS$) As Boolean
Dim p1%: p1 = InStr(pS, "{")
Dim p2%: p2 = InStr(pS, "}")
StrIsMacro = (p2 > p1 And p1 > 0)
End Function

Function StrIsNm(S) As Boolean
If Not ChrIsLetter(S) Then Exit Function
Dim J%, C$
For J = 1 To Len(S)
    C = Mid(S, J, 1)
    If Not ChrIsNmChr(C) Then Exit Function
Next
StrIsNm = True
End Function

Function StrIsSubstrInAp(S, ParamArray Ap()) As Boolean
Dim Ay(): Ay = Ap
StrIsSubstrInAp = StrIsSubstrInAy(S, Ay)
End Function

Function StrIsSubstrInAy(S, Ay) As Boolean
Dim I
For Each I In Ay
    If InStr(S, I) Then StrIsSubstrInAy = True: Exit Function
Next
End Function

Function StrLen&(S)
StrLen = Len(S)
End Function

Function StrLik(S, LikStr) As Boolean
'Debug.Print S, LikStr, S Like LikStr
StrLik = S Like LikStr
End Function

Function StrMacroAy(S, Optional ExclBkt As Boolean) As String()
Dim A$(): A = Split(S, "{")
Dim O$(), J&, P&
For J = 1 To UB(A)
    P = InStr(A(J), "}")
    If P > 1 Then
        PushNoDup O, Left(A(J), P - 1)
    End If
Next
If Not ExclBkt Then O = AyQuote(O, "{}")
StrMacroAy = O
End Function

Sub StrMacroAy__Tst()
AyAsstEq StrMacroAy("{a} b {d} {cccc}dd", ExclBkt:=True), ApSy("a", "d", "cccc")
AyAsstEq StrMacroAy("{a} b {d} {cccc}dd"), ApSy("{a}", "{d}", "{cccc}")
End Sub

Function StrNz$(S, Nz)
If S = "" Then StrNz = Nz Else StrNz = S
End Function

Function StrUnderLine$(S, Optional UnderLineChr = "-")
StrUnderLine = String(Len(S), UnderLineChr)
End Function

Sub StrWrt(S, Ft)
Dim F%: F = FtOpnOup(Ft)
Print #F, S
Close #F
End Sub

