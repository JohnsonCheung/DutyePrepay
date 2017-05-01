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
    Fny = NmBrk(FnStr)

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

Function DicNew(DicStr$, Optional BrkChr$ = "=", Optional SepChr$ = ";") As Dictionary
Dim Sy$(): Sy = Split(DicStr, SepChr)
Dim O As New Dictionary
If Not AyIsEmpty(Sy) Then
    Dim I
    For Each I In Sy
        With Brk(I, BrkChr)
            O.Add .S1, .S2
        End With
    Next
End If
Set DicNew = O
End Function

Sub DicNew__Tst()
Dim A$: A = TblCnnStr("Permit")
Dim B As Dictionary: Set B = DicNew(A)
DicBrw B
End Sub

Function DicNewLpAp(Lp$, ParamArray Ap()) As Dictionary
Dim Ay$(): Ay = Split(Lp, " ")
Dim J&, O As New Dictionary
Dim I
For Each I In Ay
    O.Add I, Ap(J)
    J = J + 1
Next
Set DicNewLpAp = O
End Function

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

Function DicSemiColonKeyStr$(A As Dictionary)
DicSemiColonKeyStr = JnSemiColon(A.Keys)
End Function

Function DicSemiColonValStr$(A As Dictionary)
DicSemiColonValStr = JnSemiColon(A.Items)
End Function

Sub DicSetRs(Dic As Dictionary, oRs As DAO.Recordset)
'Aim: Set {oRs} by {pLnFld} & {pAyV}.  Assume oRs is already .AddNew or .Edit
Const cSub$ = "Set_Rs_ByLpVv"

Dim J%, mAnFld$(): 'mAnFld = Split(pLnFld, cComma)
Dim mNmFld$, mAyV()
'mAyV = pVayv
With oRs
    For J = 0 To UBound(mAnFld$)
        mNmFld = Trim(mAnFld(J))
        .Fields(mNmFld).Value = mAyV(J)
    Next
End With
End Sub

Sub DicSetRs__Tst()
If Dlt_Tbl("xx") Then Stop
If Run_Sql("Create table xx (aa Long, bb Integer, cc Date)") Then Stop
Dim mRs As DAO.Recordset
Set mRs = CurrentDb.TableDefs("xx").OpenRecordset
mRs.AddNew
'DicSetRs mRs, "aa,bb,cc", "13", 12, "2007/12/31") Then Stop ' Should have NO error
mRs.Update

mRs.AddNew
'If DicSetRs(mRs, "aa,bb,cc", 13, 12, #1/1/2007#) Then Stop ' Should have NO error
mRs.Update

mRs.AddNew
'If DicSetRs(mRs, "aa,bb,cc", "13a", 12, #1/1/2007#) Then Stop ' Should have error
mRs.Update
mRs.Close
DoCmd.OpenTable ("xx")
End Sub

Function DicSy(A As Dictionary, KVNmStr$) As String()
Dim O$()
If A.Count = 0 Then Exit Function
Dim K
For Each K In A
    Push O, FmtNm(KVNmStr, "K V", K, A(K))
Next
DicSy = O
End Function

