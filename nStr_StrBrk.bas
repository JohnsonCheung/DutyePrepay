Attribute VB_Name = "nStr_StrBrk"
Option Compare Database
Option Explicit

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

Function Brk_ColonAs_ToCaptionNm(oCaption$, oNm$, pColonAsStr$) As Boolean
'Aim: Convert "[<<Caption>>:] <<Nam>>" into oNm and oCaption
Const cSub$ = "Brk_ColonAs_ToCaptionNm"
If Brk_Str_1ForS2(oCaption, oNm, pColonAsStr, ":") Then ss.A 1: GoTo E
Exit Function
R: ss.R
E:
End Function

Function Brk_ColonAs_ToCaptionNm__Tst()
Const cSub$ = "Brk_ColonAs_ToCaptionNm_Tst"

Dim mColonAsStr$, mNm$, mCaption$, mCase As Byte
For mCase = 1 To 2
    Select Case mCase
    Case 1
        mColonAsStr = "aa: xx"
    Case 2
        mColonAsStr = "xx"
    End Select
    If Brk_ColonAs_ToCaptionNm(mCaption, mNm, mColonAsStr) Then Stop
    Debug.Print mCase
    Debug.Print ToStr_LpAp(vbLf, "mColonAsStr, mNm, mCaption", mColonAsStr, mNm, mCaption)
    Debug.Print "-----------"
Next
End Function

Function Brk_Ffn_To2Seg(oFfnn$, oExt$, pFfn$) As Boolean
Dim mPosDot As Byte: mPosDot = InStrRev(pFfn, ".")
If mPosDot = 0 Then oFfnn = pFfn: oExt = "": Exit Function
oFfnn = Left(pFfn, mPosDot - 1)
oExt = Mid(pFfn, mPosDot)
End Function

Function Brk_Ffn_To3Seg(oDir$, oFnn$, oExt$, pFfn$) As Boolean
oDir = Fct.Nam_DirNam(pFfn)
Brk_Ffn_To3Seg = Brk_Ffn_To2Seg(oFnn, oExt, Fct.Nam_FilNam(pFfn))
End Function

Function Brk_Str_To3Seg(oS1, oS2, oS3, pS$, Optional pBrkChr$ = ":", Optional pNoTrim As Boolean = False) As Boolean
Const cSub$ = "Brk_Str_To3Seg"
Dim A$
If Brk_Str_0Or2(oS1, A, pS, pBrkChr, pNoTrim) Then ss.A 1: GoTo E
Brk_Str_To3Seg = Brk_Str_0Or2(oS2, oS3, A, pBrkChr, pNoTrim)
Exit Function
R: ss.R
E:
End Function

Function Brk_Str_To4Seg(oS1, oS2, oS3, oS4, pS$, Optional pBrkChr$ = ":", Optional pNoTrim As Boolean = False) As Boolean
Const cSub$ = "Brk_Str2Seg4"
Dim A$
If Brk_Str_To3Seg(oS1, oS2, A, pS, pBrkChr, pNoTrim) Then ss.A 1: GoTo E
Brk_Str_To4Seg = Brk_Str_0Or2(oS3, oS4, A, pBrkChr, pNoTrim)
Exit Function
R: ss.R
E:
End Function

Function Brk_Str_To5Seg(oS1, oS2, oS3, oS4, oS5, pS$, Optional pBrkChr$ = ":", Optional pNoTrim As Boolean = False) As Boolean
Const cSub$ = "Brk_Str_To5Seg"
Dim A$
If Brk_Str_To4Seg(oS1, oS2, oS3, A, pS, pBrkChr, pNoTrim) Then ss.A 1: GoTo E
Brk_Str_To5Seg = Brk_Str_0Or2(oS4, oS5, A, pBrkChr, pNoTrim)
Exit Function
R: ss.R
E:
End Function

Function StrBrk(S, BrkChr$, Optional NoTrim As Boolean) As S1S2
'Aim: Brk {S} into {S1S2}  Format of pS: <oS1><pBrkChr><oS2>) with both <oS1> & <oS2> & <pBrkChr> must exist
If BrkChr = "" Then Er "BrkChr must be given"
Dim At&: At = InStr(S, BrkChr): If At = 0 Then Er "{S} must contain {BrkChr}", S, BrkChr
Dim O As S1S2
    O = StrBrkAt(S, At, Len(BrkChr))
    If Not NoTrim Then O = S1S2Trim(O)
StrBrk = O
End Function

Function StrBrk_Nm(OAy$(), pNm$, Optional pMax As Byte = 5) As Boolean
'Aim: Break Nm into Nm1,..,Nm5
Const cSub$ = "Brk_Nm"
Const cA As Byte = 65
Const cZ As Byte = 90
Dim J%, mS As Byte, mA$: mA = ""
Clr_Ays OAy
For J = 1 To Len(pNm)
    mS = Asc(Mid(pNm, J, 1))
    If cA <= mS And mS <= cZ Then
        If Len(mA) > 0 Then
            If Add_AyEle(OAy, mA) Then ss.A 1: GoTo E
            mA = ""
        End If
    End If
    mA = mA & Chr(mS)
Next
If Len(mA) > 0 Then If Add_AyEle(OAy, mA) Then ss.A 2: GoTo E
Dim N%: N = Sz(OAy)
If N > pMax Then
    mA = ""
    For J = pMax - 1 To N - 1
        mA = mA & OAy(J)
    Next
    OAy(pMax - 1) = mA
End If
ReDim Preserve OAy(pMax - 1)
Exit Function
E:
End Function

Function StrBrk_Nm__Tst()
Dim mA$: mA = "A1A2A3A4A5A6A7"
Dim mAy$(): If StrBrk_Nm(mAy, mA) Then Stop
Debug.Print mA
Debug.Print UnderlineStr(mA)
Debug.Print ToStr_Ays(mAy, , vbLf)
End Function

Function StrBrk1(S, BrkChr$, Optional NoTrim As Boolean) As S1S2
If BrkChr = "" Then Er "BrkChr must be given"
Dim At&: At = InStr(S, BrkChr)
StrBrk1 = StrBrk1At(S, At, Len(BrkChr), NoTrim)
End Function

Function StrBrk1At(S, At&, L%, Optional NoTrim As Boolean) As S1S2
Dim O As S1S2
If At = 0 Then
    O.S1 = S
Else
    O = StrBrkAt(S, At, L)
End If
If Not NoTrim Then O = S1S2Trim(O)
StrBrk1At = O
End Function

Function StrBrk1FmEnd(S, BrkChr$, Optional NoTrim As Boolean) As S1S2
If BrkChr = "" Then Er "BrkChr must be given"
Dim At&: At = InStrRev(S, BrkChr)
StrBrk1FmEnd = StrBrk1At(S, At, Len(BrkChr), NoTrim)
End Function

Function StrBrk2(S, BrkChr$, Optional NoTrim As Boolean) As S1S2
If BrkChr = "" Then Er "BrkChr must be given"
Dim At&: At = InStr(S, BrkChr)
StrBrk2 = StrBrk2At(S, At, Len(BrkChr), NoTrim)
End Function

Function StrBrk2At(S, At&, L%, Optional NoTrim As Boolean) As S1S2
Dim O As S1S2
If At = 0 Then
    O.S2 = S
Else
    O = StrBrkAt(S, At, L)
End If
If Not NoTrim Then O = S1S2Trim(O)
StrBrk2At = O
End Function

Function StrBrk2FmEnd(S, BrkChr$, Optional NoTrim As Boolean) As S1S2
If BrkChr = "" Then Er "BrkChr must be given"
Dim At&: At = InStrRev(S, BrkChr)
StrBrk2FmEnd = StrBrk2At(S, At, Len(BrkChr), NoTrim)
End Function

Function StrBrkAt(S, At&, L%) As S1S2
If At = 0 Then Er "{At} cannot be 0", At
StrBrkAt.S1 = Left(S, At - 1)
StrBrkAt.S2 = Mid(S, At + L)
End Function

Function StrBrkFmEnd(S, BrkChr$, Optional NoTrim As Boolean) As S1S2
'Aim: Brk {S} into {S1S2}  Format of pS: <oS1><pBrkChr><oS2>) with both <oS1> & <oS2> & <pBrkChr> must exist
If BrkChr = "" Then Er "BrkChr must be given"
Dim At&: At = InStrRev(S, BrkChr): If At = 0 Then Er "{S} must contain {BrkChr}", S, BrkChr
Dim O As S1S2
    O = StrBrkAt(S, At, Len(BrkChr))
    If Not NoTrim Then O = S1S2Trim(O)
StrBrkFmEnd = O
End Function
