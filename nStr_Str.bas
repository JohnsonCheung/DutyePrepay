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

Function ExpdSeq(Beg&, N&, Optional FmtStr$ = "{N}", Optional MacroStr$ = "{N}") As String()
'Aim: Build a string to repeating {pFmtStr} {pN} times from {pBeg} with separated by {pSepChr}.  {pFmtStr} has {N} as the Idx.
Dim O$()
Dim J&
For J = Beg To Beg + N - 1
    Push O, Replace(FmtStr, MacroStr, J)
Next
ExpdSeq = O
End Function

Sub ExpdSeq__Tst()
Dim mExpr$
mExpr = "ExpdSeq(0, 10, ""a{N} as xx{N}"")"
Debug.Print "================="
Debug.Print mExpr
AyDmp Eval(mExpr) ' StrExpdSeq(1, 10, "a{N} as xx{N}")
Debug.Print
mExpr = "ExpdSeq(1, 10, ""a{N} as xx{N}"")"
Debug.Print mExpr
AyDmp Eval(mExpr) ' StrExpdSeq(1, 10, "a{N} as xx{N}")
End Sub

Function ExpdSeq1$(A%, Optional B% = 2)
ExpdSeq1 = "AA" & A & B
End Function

Sub ExpdSeq1__Tst()
Dim A$(): A = Eval("ExpdSeq(1,10)")
End Sub

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

Function StrChkEq(S1, S2) As Variant()
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
Dim O()
O = ErNew("Two strings are different:")
PushAy O, ErNew(D, P)
PushAy O, ErNew(D1, L1)
PushAy O, ErNew(D2, L2)
StrChkEq = O
End Function

Sub StrChkEq__Tst()
Dim Er(): Er = StrChkEq("1;2;3", "1;2;4")
ErBrw Er
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

Function StrExpd$(S, SyOpt, Optional Sep$ = CtComma, Optional MacroStr$ = "{?}")
'Aim: Format a string with {?} by repeatly join it after substitue {?} by pAn$(0 to N)
StrExpd = Join(StrExpdToSy(S, SyOpt, MacroStr), Sep)
End Function

Sub StrExpd__Tst()
Dim Sy$()
Dim J%
For J = 0 To 3
    Push Sy, "[" & J & "]"
Next
Debug.Print StrExpd("xxxx{?}yyyyy", Sy, " and ")
Debug.Print StrExpd("Tbl{?}", "x xx xxx")
Debug.Print StrExpd("Tbl{?}", "x xx xxx", vbCrLf)

Dim mLines$:
mLines = RplVBar("Line1:lksdjflskdf sdklfj|" & _
"Line2:klsdjf{?}klsdjf}|" & _
"Line3:ksldjfslkdf")
Sy = Split("<Itm1>,<Itm2>,<Itm3>,<Itm4>", ",")
Debug.Print StrExpd(mLines, Sy, vbCrLf)
End Sub

Function StrExpdToSy(S, SyOpt, Optional MacroStr$ = "{?}") As String()
'Aim: Format a string with {?} by repeatly join it after substitue {?} by pAn$(0 to N)
Dim OSy$(): OSy = OptSy(SyOpt)
If InStr(S, MacroStr) = 0 Then
    StrExpdToSy = ApSy(S)
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
StrExpdToSy = O
End Function

Function StrHas(S, SubStr) As Boolean
StrHas = InStr(S, SubStr) > 0
End Function

Function StrIsBlank(S) As Boolean
StrIsBlank = Trim(RmvBlankChr(S)) = ""
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
Dim P1%: P1 = InStr(pS, "{")
Dim P2%: P2 = InStr(pS, "}")
StrIsMacro = (P2 > P1 And P1 > 0)
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

