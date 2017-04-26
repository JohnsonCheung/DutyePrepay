Attribute VB_Name = "nDao_LExpr"
Option Compare Database
Option Explicit
Const cMod$ = ""

Function LExpr(oLExpr$, pNmFld$, pTypSim As eTypSim, pVraw$, Optional pIsOpt As Boolean = False) As Boolean
'Aim: Build a sql condition {oLExpr}
'Prm: {pVraw}   [x] [%x-x] [x,x,x] [>x] [>=x] [<x] [<=x] [*x] [x*] [*x*] [!%x-x] [!x,x,x] [!*x] [!x*] [!*x*] [!x] (16)
'               Eq  Rge    Lst     Gt   Ge    Lt   Le    ----- Lik ----- NRge    NLst     ------ NLik ------ Ne   (12)
Const cSub$ = "LExpr"
On Error GoTo R
If pIsOpt Then If Trim(pVraw) = "" Then ss.A 1, "Not optional prm, but pVraw is empty": GoTo E
Select Case pTypSim
Case eTypSim_Num, eTypSim_Bool, eTypSim_Str, eTypSim_Dte
Case Else
     ss.A 2, "Only TypSim: N,B,S,D will handle": GoTo E
End Select
If pNmFld = "" Then ss.A 3, "pNmfld cannot be nothing": GoTo E

'Find [OpTyp,V1,V2] by {pVraw}
Dim V1$, V2$, OpTyp As eOpTyp
''(5) [Test C2 first]  [>=] [<=] [!%] [!*x] [!*x*]
Dim C2$: C2 = Left(pVraw, 2)
Select Case C2
Case ">=": OpTyp = eGe: V1 = Mid(pVraw, 3)
Case "<=": OpTyp = Ele: V1 = Mid(pVraw, 3)
Case "!%": OpTyp = eNRge: If Brk_Str_Both(V1, V2, Mid(pVraw, 3), "-") Then ss.A 4, "no - in !%xx-xx": GoTo E
Case "!*": OpTyp = eNLik: V1 = Mid(pVraw, 2)
Case Else
    ''(9) [Test C1 next]   [%]  [>]  [<]  [*x]  [*x*]  [!x:x:x] [!x*] [!x]
    Dim C1$: C1 = Left(C2, 1)
    Select Case C1
    Case "%": OpTyp = eRge: If Brk_Str_Both(V1, V2, Mid(pVraw, 2), "-") Then ss.A 5, "no - in %xx-xx": GoTo E
    Case ">": OpTyp = eGt: V1 = Mid(pVraw, 2)
    Case "<": OpTyp = eLt: V1 = Mid(pVraw, 2)
    Case "*": OpTyp = eLik: V1 = pVraw
    Case "!": V1 = Mid(pVraw, 2) '[!x,x,x] [!x*] [!x]
                If InStr(V1, CtComma) > 0 Then
                    OpTyp = eNLst
                Else
                    If InStr(V1, "*") > 0 Then
                        OpTyp = eNLik
                    Else
                        OpTyp = eNe
                    End If
                End If
    Case Else
        ''(1) [Test x,x,x]
        Dim Pos As Byte: Pos = InStr(pVraw, CtComma)
        If Pos > 0 Then
            OpTyp = eLst: V1 = pVraw
        Else
            ''(1) [x]
            OpTyp = eEq: V1 = pVraw
        End If
    End Select
End Select

'[Find Q]
''All Op: Eq , Rge, Lst, Gt, Ge, Lt, Le, Lik, NRge, NLst, NLik, Ne
Dim Q$
Select Case pTypSim
''[Bool: Op=(Eq,Ne)]
Case eTypSim_Bool
    Select Case OpTyp
    Case eOpTyp.eEq, eOpTyp.eNe
    Case Else:  ss.A 6, "Boolean data only allow = or <>", , "OpTyp", OpTyp: GoTo E
    End Select
    If V1 <> "True" And V1 <> "False" Then ss.A 7, "Boolean data allow value True or False", , "V1", V1: GoTo E
Case eTypSim_Dte
''[Dte: Op=(!Lik,NLik)]
    Select Case OpTyp
    Case eOpTyp.eLik, eOpTyp.eNLik:  ss.A 8, "Date data not allow Like or not like", , "OpTyp", OpTyp: GoTo E
    Case Else
    End Select
    Q = "#"
Case eTypSim_Str
    Q = CtSngQ
End Select

'[Reject Some Value for some TypSim]
''All Op: Eq , Rge, Lst, Gt, Ge, Lt, Le, Lik, NRge, NLst, NLik, Ne
Dim Ay$(), J%
Select Case OpTyp
Case eOpTyp.eEq, eOpTyp.eGt, eOpTyp.eGe, eOpTyp.eLt, eOpTyp.Ele, eOpTyp.eLik, eOpTyp.eNLik, eOpTyp.eNe
    If Cv_Vraw2Val(V1, V1, pTypSim) Then ss.A 9, "Field 1 has invalid value": GoTo E
Case eOpTyp.eRge, eOpTyp.eNRge
    Dim mV1: If Cv_Vraw2Val(mV1, V1, pTypSim) Then ss.A 10, "Field 1 has invalid value", , "V1", V1: GoTo E
    Dim mV2: If Cv_Vraw2Val(mV2, V2, pTypSim) Then ss.A 11, "Field 2 has invalid value", , "V1", V2: GoTo E
    If mV1 > mV2 Then ss.A 12, "Field 2 > Field 1 for Between or Not Between", , "V1,V2", V1, V2: GoTo E
Case eOpTyp.eLst, eOpTyp.eNLst
    Ay = Split(V1, CtComma)
    For J = LBound(Ay) To UBound(Ay)
        If Cv_Vraw2Val(Ay(J), Ay(J), pTypSim) Then ss.A 13, "Some field of list data has invalid value", "Ay(J)", Ay(J): GoTo E
    Next
Case Else
    If Cv_Vraw2Val(Ay(J), Ay(J), pTypSim) Then ss.A 1: GoTo E
End Select

'[Normalize V1,V2 if Q<>'']
If Q <> "" Then
    Select Case OpTyp
    Case eOpTyp.eEq, eOpTyp.eGt, eOpTyp.eGe, eOpTyp.eLt, eOpTyp.Ele, eOpTyp.eLik, eOpTyp.eNLik, eOpTyp.eNe
        V1 = Q & V1 & Q
    Case eOpTyp.eRge, eOpTyp.eNRge
        V1 = Q & V1 & Q: V2 = Q & V2 & Q
    Case eOpTyp.eLst, eOpTyp.eNLst
        Ay = Split(V1, CtComma)
        For J = LBound(Ay) To UBound(Ay)
            Ay(J) = Q & Ay(J) & Q
        Next
        V1 = Join(Ay, CtComma)
    Case Else
         ss.A 14, "Unexpected OpTyp for data need quote", , "Q,OpTyp", Q, OpTyp: GoTo E
    End Select
End If

X:
Static AyOpFmtStr$(1 To 12)
If AyOpFmtStr(1) = "" Then
    AyOpFmtStr(eOpTyp.eLik) = "({0} like {1})"
    AyOpFmtStr(eOpTyp.eEq) = "({0}={1})"
    AyOpFmtStr(eOpTyp.eGe) = "({0}>={1})"
    AyOpFmtStr(eOpTyp.eGt) = "({0}>{1})"
    AyOpFmtStr(eOpTyp.Ele) = "({0}<={1})"
    AyOpFmtStr(eOpTyp.eLst) = "({0} in ({1}))"
    AyOpFmtStr(eOpTyp.eLt) = "({0}<{1})"
    AyOpFmtStr(eOpTyp.eNLik) = "({0} not like {1})"
    AyOpFmtStr(eOpTyp.eNLst) = "({0} not in ({1}))"
    AyOpFmtStr(eOpTyp.eNRge) = "({1}>{0} or {0}>{2})"
    AyOpFmtStr(eOpTyp.eRge) = "({0} between {1} and {2})"
    AyOpFmtStr(eOpTyp.eNe) = "({0}<>{1})"
End If
oLExpr = Fmt_Str(AyOpFmtStr(OpTyp), pNmFld, V1, V2)
Exit Function
R: ss.R
E: LExpr = True: ss.B cSub, cMod, "pIsOpt,pNmfld,pTypSim,pVraw", pIsOpt, pNmFld, pTypSim, pVraw
End Function

Function LExpr__Tst()
Const cSub$ = "LExpr_Tst"
Dim J%, N%, mCndn$
N = 20
ReDim mN$(N)
J = -1
J = J + 1: mN(J) = "!"
J = J + 1: mN(J) = "!123"
J = J + 1: mN(J) = "!123,124"
J = J + 1: mN(J) = "!%123-124"
J = J + 1: mN(J) = "!123-124"
J = J + 1: mN(J) = "!*123"
J = J + 1: mN(J) = "!*123-1234"
J = J + 1: mN(J) = "123:124"
J = J + 1: mN(J) = "123-124"
J = J + 1: mN(J) = "%123:124"
J = J + 1: mN(J) = "%123-124"
J = J + 1: mN(J) = "%125-124"
J = J + 1: mN(J) = "123,124"
J = J + 1: mN(J) = "124"
J = J + 1: mN(J) = "*123"
J = J + 1: mN(J) = "123*"
J = J + 1: mN(J) = "*123*"
J = J + 1: mN(J) = "!*123"
J = J + 1: mN(J) = "!123*"
J = J + 1: mN(J) = "!*123"
Shw_DbgWin
Set_Silent
For J = 0 To N
    If mN(J) = "" Then Exit For
    Debug.Print J, mN(J),
    If LExpr(mCndn, "abc", eTypSim_Str, mN(J)) Then
        Debug.Print "<-- Error"
    Else
        Debug.Print "<-- "; mCndn
    End If
Next
For J = 0 To N
    If mN(J) = "" Then Exit For
    Debug.Print J, mN(J),
    If LExpr(mCndn, "abc", eTypSim_Num, mN(J)) Then
        Debug.Print "<-- Error"
    Else
        Debug.Print "<-- "; mCndn
    End If
Next
Shw_DbgWin
Set_Silent_Rst
Exit Function
E:
End Function

Function LExpr_ByAyNm2V(oLExpr$, pAyNm2V() As tNm2V, Optional pAlwNull As Boolean = False) As Boolean
'Aim: Build {oLExpr} by NEW value in {pAyNm2V}.  Any Null triggers error.
Const cSub$ = "LExpr_ByAyNm2V"
oLExpr = ""
Dim J%, mIsEq As Boolean
For J = 0 To Siz_An2V(pAyNm2V) - 1
    With pAyNm2V(J)
        If VarType(.NewV) = vbNull Then
            If Not pAlwNull Then ss.A 1, "The one of the element of .NewV in pAyNm2V is Null", , "The Ele,J", pAyNm2V(J).Nm, J: GoTo E
            oLExpr = Add_Str(oLExpr, Q_S(.Nm, "IsNull(*)"), " and ")
        Else
            oLExpr = Add_Str(oLExpr, .Nm & "=" & Q_V(.NewV), " and ")
        End If
    End With
Next
Exit Function
E: LExpr_ByAyNm2V = True: ss.B cSub, cMod, "pAyNm2V", ToStr_AyNm2V(pAyNm2V)
End Function

Function LExpr_ByLpAp(oLExpr$, pLp$, ParamArray pAp()) As Boolean
'Aim: Build condition {oLExpr} by {pLn} and values in {pAp()}
Const cSub$ = "LExpr_ByLpAp"
If LExpr_ByLpVv(oLExpr, pLp, CVar(pAp)) Then ss.A 1: GoTo E
Exit Function
R: ss.R
E: LExpr_ByLpAp = True: ss.B cSub, cMod, "pLp,pAp", pLp, ToStr_Vayv(CVar(pAp))
End Function

Function LExpr_ByLpVv(oLExpr$, pLp$, pVayv) As Boolean
'Aim: Build condition {oLExpr} by {pLn} and a variant which is array of variant of value {pVayv}.  {Vayv} mean Variant that storing array of variant. {Val} means value
Const cSub$ = "LExpr_ByLpVv"
If VarType(pVayv) <> vbArray + vbVariant Then ss.A 1, "VarType of pVayv must be Array+Var", , "VarType(pVayv)", VarType(pVayv): GoTo E
oLExpr = ""
Dim mAn$(): mAn = Split(pLp, CtComma)
Dim mAyV(): mAyV = pVayv
Dim N1%: N1 = Sz(mAn)
Dim N2%: N2 = Sz(mAyV)
If N1 <> N2 Then ss.A 1, "Cnt in pLn & pV() not match", , "N1,N2", N1, N2: GoTo E
Dim J%: For J = 0 To N1 - 1
    Dim mA$: If Join_NmV(mA, mAn(J), mAyV(J)) Then ss.A 2: GoTo E
    oLExpr = Add_Str(oLExpr, mA, " and ")
Next
Exit Function
R: ss.R
E: LExpr_ByLpVv = True: ss.B cSub, cMod, "pLn,pV", pLp, ToStr_Vayv(pVayv)
End Function

Function LExpr_ByLpVv__Tst()
Const cSub$ = "LExpr_ByLpAp_Tst"
Dim mCndn$, mLn$, mV1, mV2, mV3
Dim mRslt As Boolean, mCase As Byte: mCase = 1
Select Case mCase
Case 1
    mLn$ = "ss,nn,dd"
    mV1 = "aa"
    mV2 = 11
    mV3 = #2/1/2007#
End Select
mRslt = LExpr_ByLpAp(mCndn, mLn, mV1, mV2, mV3)
Shw_Dbg cSub, cMod, "mRslt,mCndn,mLn,mV1,mV2,mV3", mRslt, mCndn, mLn, mV1, mV2, mV3
End Function

Function LExpr_InFrm(oLExpr$, pFrm As Form, pLmPk$) As Boolean
'Aim: Build {oLExpr} by OldValue of {pLmPk$} in {pFrm} with optional to replace the variable name by {pLnNew}
Const cSub$ = "LExpr_InFrm"
Dim mAyNm2V() As tNm2V: If Fnd_An2V_ByFrm(mAyNm2V, pFrm, pLmPk) Then ss.A 1: GoTo E
LExpr_InFrm = LExpr_ByAyNm2V(oLExpr, mAyNm2V)
Exit Function
R: ss.R
E: LExpr_InFrm = True: ss.B cSub, cMod, ""
End Function
