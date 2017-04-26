Attribute VB_Name = "nDta_Dic"
Option Compare Database
Option Explicit

Sub DicAsg(A As Dictionary, FnStr$, O0 _
    , Optional O1 _
    , Optional O2 _
    , Optional O3 _
    , Optional O4 _
    , Optional O5 _
    , Optional O6 _
    , Optional O7 _
    , Optional O8 _
    , Optional O9 _
    , Optional O10 _
    , Optional O11 _
    , Optional O12 _
    , Optional O13 _
    , Optional O14 _
    , Optional O15 _
    )
Dim Fny$()
    Fny = NmstrBrk(FnStr)

Dim K, J%, V
For Each K In Fny
    V = A(K)
    Select Case J
    Case 0: O0 = V
    Case 1: O1 = V
    Case 2: O2 = V
    Case 3: O3 = V
    Case 4: O4 = V
    Case 5: O5 = V
    Case 6: O6 = V
    Case 7: O7 = V
    Case 8: O8 = V
    Case 9: O9 = V
    Case 10: O10 = V
    Case 11: O11 = V
    Case 12: O12 = V
    Case 13: O13 = V
    Case 14: O14 = V
    Case 15: O15 = V
    End Select
Next
End Sub

Sub DicBrw(A As Dictionary)
DtBrw DicDt(A), "Dic"
End Sub

Function DicChkEq(D1 As Dictionary, D2 As Dictionary) As Variant()
Dim O()
If D1.Count <> D2.Count Then
    Push O, FmtQQ("Count is diff: [?] [?]", D1.Count, D2.Count)
Else
    O = AyChkSam(D1.Keys, D2.Keys)
    Dim K
    If AyIsEmpty(O) Then
        For Each K In D1.Keys
            If D1(K) <> D2(K) Then
                Push O, FmtQQ("Values of Key[?} are dif: [?] / [?]", K, D1(K), D2(K))
            End If
        Next
    End If
End If
If Not AyIsEmpty(O) Then
    Push O, "Two Dic's key not match"
End If
DicChkEq = O
End Function

Sub DicDmp(A As Dictionary)
Dim J&
Dim K
For Each K In A.Keys
    Debug.Print J & " [" & K & "] = [" & A(K) & "]"
    J = J + 1
Next
End Sub

Function DicDt(A As Dictionary) As Dt
Dim D()
If Not DicIsEmpty(A) Then
    Dim I, V
    For Each I In A
        V = A(I)
        Push D, Array(I, V, VbTyStr(VarType(V)))
    Next
End If
DicDt = DtNew(LvsSplit("Itm Val VbTy"), D)
End Function

Function DicIsEmpty(A As Dictionary) As Boolean
DicIsEmpty = A.Count = 0
End Function

Function DicIsEq(D1 As Dictionary, D2 As Dictionary) As Boolean
DicIsEq = Sz(DicChkEq(D1, D2)) = 0
End Function

Function DicNew(DicStr$) As Dictionary
Dim A$(): A = AyRmvBlank(SclSy(DicStr))
Dim O As New Dictionary
If AyIsEmpty(A) Then Set DicNew = O: Exit Function
Dim I
For Each I In A
    With StrBrk(I, "=")
        O.Add .S1, .S2
    End With
Next
Set DicNew = O
End Function

Sub DicNew__Tst()
Dim A$: A = TblCnnStr("Permit")
Dim B As Dictionary: Set B = DicNew(A)
DicBrw B
End Sub

Function DicNewSyAp(OptSy, ParamArray Ap()) As Dictionary
Dim O As New Dictionary
'Dim mAn$(): mAn = Split(pLp, CtComma)
'Dim mAyV(): mAyV = Ay
'Dim N1%: N1 = Sz(mAn)
'Dim N2%: N2 = Sz(mAyV)
'If N1 <> N2 Then ss.A 1, "Cnt in pAp & pLp are diff", , "Cnt in pAp,Cnt in pLp", N2, N1: GoTo E
'ReDim mAm(N1 - 1) As tMap
'If Set_Am_F1(mAm, mAn) Then ss.A 2: GoTo E
'Dim J%, mA$
'For J = 0 To N1 - 1
'    If Not IsMissing(mAyV(J)) Then
'        If (VarType(mAyV(J)) And vbArray) Then
'            mAm(J).F2 = Join(mAyV(J), ",")
'        Else
'            mAm(J).F2 = mAyV(J)
'        End If
'    End If
'Next
Set DicNewSyAp = O
End Function

Function Get_Am_ByLm(pLm$, Optional pBrkChr$ = "=", Optional pSepChr$ = CtSemiColon) As tMap()
Dim mAmStr$(): mAmStr = Split(pLm, pSepChr$)
Dim NMap%: NMap = Sz(mAmStr)
If NMap = 0 Then Exit Function
ReDim mAm(0 To NMap - 1) As tMap
Dim I%: For I = 0 To NMap - 1
    If Brk_Str2Map(mAm(I), mAmStr(I), pBrkChr$) Then Exit Function
Next
Get_Am_ByLm = mAm
End Function

Function Get_Am_ByLpAp(pLp$, ParamArray pAp()) As tMap()
Get_Am_ByLpAp = Get_Am_ByLpVv(pLp, CVar(pAp))
End Function

