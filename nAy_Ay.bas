Attribute VB_Name = "nAy_Ay"
Option Compare Database
Option Explicit
Private Type DifIdx
    DifIdx() As Long
    IsMore As Boolean
End Type

Function AyAdd(Ay1, Ay2)
Dim O: O = Ay1
Dim U&: U = UB(Ay2)
If U >= 0 Then
    Dim I
    For Each I In Ay2
        Push O, I
    Next
End If
AyAdd = O
End Function

Function AyAddCol_ConstAft(Ay, ConstCol) As Variant()
Dim O(), U&, J&
For J = 0 To UB(Ay)
    Push O, Array(Ay(J), ConstCol)
Next
AyAddCol_ConstAft = O
End Function

Sub AyAddCol_ConstAft__Tst()
Dim Ay$(): Ay = LvsSplit("a b c d")
Dim Act(): Act = AyAddCol_ConstAft(Ay, "2")
Dim Dr()
Debug.Assert Sz(Act) = 4
Dr = Act(0): Debug.Assert Dr(0) = "a": Debug.Assert Dr(1) = "2"
Dr = Act(1): Debug.Assert Dr(0) = "b": Debug.Assert Dr(1) = "2"
Dr = Act(2): Debug.Assert Dr(0) = "c": Debug.Assert Dr(1) = "2"
Dr = Act(3): Debug.Assert Dr(0) = "d": Debug.Assert Dr(1) = "2"
End Sub

Function AyAddCol_ConstBef(Ay, ConstCol) As Variant()
Dim O(), U&, J&
For J = 0 To UB(Ay)
    Push O, Array(ConstCol, Ay(J))
Next
AyAddCol_ConstBef = O
End Function

Sub AyAddCol_ConstBef__Tst()
Dim Ay$(): Ay = LvsSplit("a b c d")
Dim Act(): Act = AyAddCol_ConstBef(Ay, "2")
Dim Dr()
Debug.Assert Sz(Act) = 4
Dr = Act(0): Debug.Assert Dr(0) = "2": Debug.Assert Dr(1) = "a"
Dr = Act(1): Debug.Assert Dr(0) = "2": Debug.Assert Dr(1) = "b"
Dr = Act(2): Debug.Assert Dr(0) = "2": Debug.Assert Dr(1) = "c"
Dr = Act(3): Debug.Assert Dr(0) = "2": Debug.Assert Dr(1) = "d"
End Sub

Function AyAddPfx(Ay, Pfx) As String()
Dim U&: U = UB(Ay)
Dim O$(): ReSz O, U
Dim J&
For J = 0 To U
    O(J) = Pfx & Ay(J)
Next
AyAddPfx = O
End Function

Function AyAddPfxSfx(Ay, Pfx, Sfx) As String()
Dim U&: U = UB(Ay)
Dim O$(): ReSz O, U
Dim J&
For J = 0 To U
    O(J) = Pfx & Ay(J) & Sfx
Next
AyAddPfxSfx = O
End Function

Function AyAddSfx(Ay, Sfx) As String()
Dim U&: U = UB(Ay)
Dim O$(): ReSz O, U
Dim J&
For J = 0 To U
    O(J) = Ay(J) & Sfx
Next
AyAddSfx = O
End Function

Sub AyAsg(Ay, ParamArray OAp())
Dim J%, Av()
Av = OAp
For J = 0 To Min(UB(Av), UB(Ay))
    VarAsg Ay(J), OAp(J)
Next
End Sub

Sub AyAsgSubAyIdx(Ay, SubAy, ParamArray OAp())
Dim U%: U = UB(SubAy)
Dim J%
For J = 0 To U
    OAp(J) = AyIdx(Ay, SubAy(J))
Next
End Sub

Sub AyAsstDupEle(Ay, ParamArray MsgAp())
Dim Av(): Av = MsgAp
ErAsst AyChkDupEle(Ay), Av
End Sub

Sub AyAsstDupEle__Tst()
Dim Ay: Ay = Array(1, 2, 3, 2)
AyAsstDupEle Ay ' , "AyAsstDupEle__Tst"
End Sub

Sub AyAsstEq(Ay1, Ay2, ParamArray MsgAp())
Dim Av(): Av = MsgAp
ErAsst AyChkEq(Ay1, Ay2), Av
End Sub

Sub AyAsstEqExa(Ay1, Ay2, ParamArray MsgAp())
Dim Av(): Av = MsgAp
ErAsst AyChkEqExa(Ay1, Ay2), Av
End Sub

Sub AyAsstEqSz(Ay1, Ay2, ParamArray MsgAp())
Dim Av(): Av = MsgAp
ErAsst AyChkEqSz(Ay1, Ay2), Av
End Sub

Sub AyAsstSamSet(Ay1, Ay2, ParamArray MsgAp())
Dim Av(): Av = MsgAp
ErAsst AyChkSamSet(Ay1, Ay2), Av
End Sub

Sub AyAsstZerOrPos(Ay, ParamArray MsgAp())
Dim Av(): Av = MsgAp
ErAsst AyChkZerOrPos(Ay), Av
End Sub

Sub AyBrw(Ay, Optional Pfx$ = "Ay", Optional WithIdx As Boolean)
If WithIdx Then
    DrAyBrw AyDrAy(Ay)
Else
    Dim F$: F = TmpFt(Pfx)
    AyWrt Ay, F
    FtBrw F, True
End If
End Sub

Function AyChkDupEle(Ay) As Dt
Dim A As Dt: A = DtWhere(AyDistDt(Ay), "Cnt", ">1")

Dim N&: N = DtNRec(A)
Dim ODt As Dt
If N > 0 Then
    Dim J&, V(), C&()
    ODt = ErNew("Given Ay has {N}-Element.  {M} of are duplicated", Sz(Ay), N)
    For J = 0 To N - 1
        ODt = ErApd(ODt, ".EleVal-{V} are found {N}-times", A.DrAy(J)(0), A.DrAy(J)(1))
    Next
End If
AyChkDupEle = ODt
End Function

Sub AyChkDupEle__Tst()
Dim Ay: Ay = Array(1, 1, 1, 2, 2, 3, 1, "a")
DtBrw AyChkDupEle(Ay)
End Sub

Function AyChkEq(Ay1, Ay2) As Dt
AyChkEq = AyChkEqEr(Ay1, Ay2, False)
End Function

Sub AyChkEq__Tst()
Debug.Assert ErIsSom(AyChkEq(Array(1, 2), Array("1", 2))) = True
Debug.Assert ErIsSom(AyChkEq(Array(1, 1), Array("1", 2))) = True
Debug.Assert ErIsSom(AyChkEq(Array("1", 2), Array("1", 2))) = False
Debug.Assert ErIsSom(AyChkEq(Array(CByte(1), 2), Array(CInt(1), 2))) = False

Dim Er As Dt:
Dim Ay1:
Dim Ay2:

Ay2 = Array(1, 2, 3, 4, 6, 5, 7)
Ay1 = Array(1, 2, 3, 4, 57, 4, 1)
Er = AyChkEq(Ay1, Ay2)
'Debug.Assert ErIsSom(Er)
DtBrw Er

Ay1 = Array(1, 2, 3, 4, 5)
Ay2 = Array(1, 2, 3, 4, CByte(5))
Er = AyChkEq(Ay1, Ay2)
Debug.Assert Not ErIsSom(Er)


End Sub

Function AyChkEqExa(Ay1, Ay2) As Dt
AyChkEqExa = AyChkEqEr(Ay1, Ay2, True)
End Function

Sub AyChkEqExa__Tst()
Debug.Assert ErIsSom(AyChkEqExa(Array(1, 2), Array("1", 2))) = True
Debug.Assert ErIsSom(AyChkEqExa(Array(1, 1), Array("1", 2))) = True
Debug.Assert ErIsSom(AyChkEqExa(Array("1", 2), Array("1", 2))) = False
Debug.Assert ErIsSom(AyChkEqExa(Array(CByte(1), 2), Array(CInt(1), 2))) = True

Dim D As Dt:
Dim Ay1
Dim Ay2

Ay2 = Array(1, 2, 3, 4, 6)
Ay1 = Array(1, 2, 3, 4, 5)
D = AyChkEqExa(Ay1, Ay2)
Debug.Assert ErIsSom(D)

Ay1 = Array(1, 2, 3, 4, 5)
Ay2 = Array(1, 2, 3, 4, CByte(5))
D = AyChkEqExa(Ay1, Ay2)
Debug.Assert ErIsSom(D)

Ay1 = Array(1, 2, 3, 4, CByte(5))
Ay2 = Array(1, 2, 3, 4, CByte(5))
D = AyChkEqExa(Ay1, Ay2)
Debug.Assert Not ErIsSom(D)
End Sub

Function AyChkEqSz(Ay1, Ay2) As Dt
Dim UU&, U1&, U2&, U&, J&, O As Dt
U1 = UB(Ay1)
U2 = UB(Ay2)
If U1 = U2 Then Exit Function

U = Max(U1, U2)
UU = Min(9, U)
O = ErNew("Two Ay of {U1} and {U2} are dif sz:", U1, U2)
O = ErApd(O, ".First {U}-Ele of Ay1 and Ay2", UU)
Dim V1: V1 = AyFstUEle(Ay1, UU, True)
Dim V2: V2 = AyFstUEle(Ay2, UU, True)
For J = 0 To UU
    O = ErApd(O, "..", J, V1(J), V2(J))
Next
AyChkEqSz = O
End Function

Sub AyChkEqSz__Tst()
DtBrw AyChkEqSz(Array(1, 2), Array(1, 2, 3))
End Sub

Function AyChkSamSet(Ay1, Ay2) As Dt
End Function

Function AyChkZerOrPos(Ay) As Dt
Dim ErI&(), J&
For J = 0 To UB(Ay)
    If Ay(J) < 0 Then
        Push ErI, J
    End If
Next
If AyIsEmpty(ErI) Then Exit Function
AyChkZerOrPos = ErNew("Given {Ay{-of-{U} has {U-Ele} being negative", AyJn(Ay, " "), UB(Ay), UB(ErI))
End Function

Function AyCoverIdx%(QtyAy, OH)
'Aim: Return Idx so that SumOf(QtyAy, for 0 to Idx) is just > OH
Dim J&, Q
Q = 0
For J = 0 To UB(QtyAy)
    Q = Q + QtyAy(J)
    If Q >= OH Then AyCoverIdx = J: Exit Function
Next
AyCoverIdx = -1
End Function

Function AyCut(Ay, At&) As Variant()
'Aim: Split Ay$() into 2, first {At} in {O1} and rest in {O2}
Dim O1, O2
    O1 = Ay
    Erase O1
    O2 = O1
    
    Dim N%: N = Siz_Ay(Ay)
    If At >= N Then
        O2 = Ay
    ElseIf At <= 0 Then
        O1 = Ay
    Else
        ReDim O1(At - 1)
        ReDim O2(N - At - 1)
        Dim J&
        For J = 0 To At - 1
            O1(J) = Ay(J)
        Next
        For J = 0 To N - At - 1
            O2(J) = Ay(J + At)
        Next
    End If
AyCut = Array(O1, O2)
End Function

Function AyCut__Tst()
Dim mA$(5), J&
For J = 0 To 5: mA(J) = J: Next
For J = 0 To 6
    Dim Act: Act = AyCut(mA, J)
    Dim A1, A2
    A1 = Act(0)
    A2 = Act(1)
    Debug.Print "Splitting " & J & "...."
    Debug.Print Join(A1, CtComma) & "<---"
    Debug.Print Join(A2, CtComma) & "<---"
Next
End Function

Function AyDistDt(Ay) As Dt
Dim OF$()
OF = ApSy("Val", "Cnt")
Dim OD()
    Dim V(), C&()
    If Sz(Ay) > 0 Then
        Dim I, J&, Fnd As Boolean
        For Each I In Ay
            Fnd = False
            For J = 0 To UB(V)
                If I = V(J) Then
                    C(J) = C(J) + 1
                    GoTo Nxt
                End If
            Next
            Push V, I
            Push C, 1
Nxt:
        Next
        Dim Dr()
        Dim U&: U = UB(V)
        ReDim OD(U)
        For J = 0 To U
            Dr = Array(V(J), C(J))
            OD(J) = Dr
        Next
    End If
AyDistDt = DtNew(OF, OD)
End Function

Sub AyDistDt__Tst()
Dim Ay: Ay = Array(1, 1, 1, 2, 2, 3, 1, "a")
DtBrw AyDistDt(Ay)
End Sub

Sub AyDmp(Ay)
Dim J&
For J = 0 To UB(Ay)
    Debug.Print Ay(J)
Next
End Sub

Function AyDrAy(Ay) As Variant()
Dim U&: U = UB(Ay)
Dim O(), J&
ReSz O, U
For J = 0 To U
    O(J) = Array(Ay(J))
Next
AyDrAy = O
End Function

Function AyDupItm(Ay)
If AyIsEmpty(Ay) Then AyDupItm = Ay: Exit Function
Dim O: O = Ay: Erase O
Dim A: A = O
Dim I
For Each I In Ay
    If AyHas(A, I) Then
        Push O, I
    Else
        Push A, I
    End If
Next
AyDupItm = O
End Function

Sub AyDupItm__Tst()
Dim Act, Exp
Act = AyDupItm(Array(1, 2, 3, 1, 3))
Exp = Array(1, 4)
Dim Er As Dt
Er = AyChkEq(Act, Exp)
Er = ErExplain(Er, "AyDupImt_Tst fail")
ErAsst Er
End Sub

Sub AyEachEle(Ay, Fn$, ParamArray Ap())
If AyIsEmpty(Ay) Then Exit Sub
Dim I, Av()
Av = Ap
Av = AyInsAt(Av, 0, 0)
For Each I In Ay
    VarAsg I, Av(0)
    RunAv Fn, Av
Next
End Sub

Function AyExcl(Ay, Fct$, ParamArray Ap())
If AyIsEmpty(Ay) Then AyExcl = Ay: Exit Function
Dim O: O = Ay: Erase O
Dim I
Dim Av(): Av = Ap: Av = AyInsAt(Av)
For Each I In Ay
    Av(0) = I
    If Not RunAv(Fct, Av) Then Push O, I
Next
AyExcl = O
End Function

Function AyExpandAy(Itm_or_Ay)
Dim J&, O()
For J = 0 To UB(Itm_or_Ay)
    If IsArray(Itm_or_Ay(J)) Then
        PushAy O, Itm_or_Ay(J)
    Else
        Push O, Itm_or_Ay(J)
    End If
Next
AyExpandAy = O
End Function

Function AyFstUEle(Ay, Optional U& = 0, Optional IsExtToU As Boolean)
Dim O: O = Ay: Erase O
Dim U1&: U1 = UB(Ay)
Dim UU&: UU = Min(U, UB(Ay))
ReSz O, UU
Dim J&
For J = 0 To UU
    O(J) = Ay(J)
Next
If IsExtToU Then
    If U > U1 Then ReDim Preserve O(U)
End If
AyFstUEle = O
End Function

Function AyGpByPfx(Ay, Pfx$, Optional IsNoPfxInFstEle As Boolean) As Variant()
If AyIsEmpty(Ay) Then Exit Function
If Not IsNoPfxInFstEle Then If IsPfx(Ay(0), Pfx) Then Er "FstEle of {Ay} does not have {Pfx}", Ay(0), Pfx
Dim O()
Dim B&
    Dim M
    M = Ay
    Erase M
Dim J&
Dim NoMore As Boolean
For J = 0 To UB(Ay)
    If IsPfx(Ay(J), Pfx) Then
        If Not AyIsEmpty(M) Then Push O, M
        Erase M
        Push M, Ay(J)
    End If
Next
If Not AyIsEmpty(M) Then Push O, M
AyGpByPfx = O
End Function

Function AyGpBySamVal(Ay, Optional ExclSingleEleGp As Boolean) As Variant()
Dim U&: U = UB(Ay): If U = -1 Then Exit Function
Dim Las: Las = Ay(0)
Dim OO()
    Dim G&()
    Dim Cur, R&
    ReDim G(1): G(0) = 0
    For R = 1 To U
        Cur = Ay(R)
        If Las <> Cur Then
            Las = Cur
            G(1) = R - 1
            Push OO, G
            ReDim G(1)
            G(0) = R
        End If
    Next
    G(1) = U
    Push OO, G

Dim O()
    If ExclSingleEleGp Then
        Dim J&
        For J = 0 To UB(OO)
            G = OO(J)
            If G(1) > G(0) Then Push O, G
        Next
    Else
        O = OO
    End If
AyGpBySamVal = O
End Function

Function AyHas(Ay, I) As Boolean
If AyIsEmpty(Ay) Then Exit Function
Dim Itm
For Each Itm In Ay
    If Itm = I Then AyHas = True: Exit Function
Next
End Function

Function AyHasDup(Ay) As Boolean
Dim U&: U = UB(Ay): If U <= 0 Then Exit Function
Dim A: A = Ay: Erase A
Push A, Ay(0)
Dim J&
For J = 1 To U
    If AyHas(A, Ay(J)) Then AyHasDup = True: Exit Function
    Push A, Ay(J)
Next
End Function

Function AyIdx&(Ay, I)
Dim U&
    U = UB(Ay)
If U >= 0 Then
    Dim J&
    For J = 0 To U
        If Ay(J) = I Then AyIdx = J: Exit Function
    Next
End If
AyIdx = -1
End Function

Function AyIdxAy(Ay, SubAy) As Long()
If AyIsEmpty(SubAy) Then Exit Function
Dim U&: U = UB(SubAy)
Dim O&(): ReSz O, U
Dim J&
For J = 0 To U
    O(J) = AyIdx(Ay, SubAy(J))
Next
AyIdxAy = O
End Function

Function AyIdxDic(Ay) As Dictionary
Dim O As New Dictionary
If AyIsEmpty(Ay) Then GoTo X
Dim V, I&
For Each V In Ay
    O.Add V, I
    I = I + 1
Next
X:
Set AyIdxDic = O
End Function

Function AyInsAt(Ay, Optional At& = 0, Optional I)
Dim N&: N = Sz(Ay)
Dim O: O = Ay
ReDim Preserve O(N)
Dim J&
For J = N To At + 1 Step -1
    VarAsg Ay(J - 1), O(J)
Next
VarAsg I, O(At)
AyInsAt = O
End Function

Function AyIntersect(Ay1, Ay2)
Dim O: O = Ay1: Erase O
Dim U1&, U2&: U1 = UB(Ay1): U2 = UB(Ay2)
If U1 = -1 Or U2 = -1 Then AyIntersect = O: Exit Function
Dim J&
For J = 0 To U1
    If AyHas(Ay2, Ay1(J)) Then Push O, Ay1(J)
Next
AyIntersect = O
End Function

Function AyIsAllEmptyEle(Dr) As Boolean
If AyIsEmpty(Dr) Then Exit Function
Dim V
For Each V In Dr
    If Not IsEmpty(V) Then Exit Function
Next
AyIsAllEmptyEle = True
End Function

Function AyIsEmpty(Ay) As Boolean
AyIsEmpty = Sz(Ay) = 0
End Function

Function AyIsEq(Ay1, Ay2) As Boolean
If VarType(Ay1) <> VarType(Ay2) Then Exit Function
Dim U&: U = UB(Ay1)
If U <> UB(Ay2) Then Exit Function
Dim J&
For J = 0 To U
    If Not VarIsEq(Ay1(J), Ay2(J)) Then Exit Function
Next
AyIsEq = True
End Function

Function AyIsSam(Ay1, Ay2) As Boolean

End Function

Function AyJn$(Ay, Optional Sep$)
AyJn = Join(AySy(Ay), Sep)
End Function

Function AyJnComma$(Ay)
AyJnComma = Join(AySy(Ay), CtComma)
End Function

Function AyJnScl$(Ay)
AyJnScl = AyJn(Ay, ";")
End Function

Function AyJnSpc$(Ay)
AyJnSpc = Join(AySy(Ay), " ")
End Function

Function AyLasEle(Ay)
If AyIsEmpty(Ay) Then Er "AyLasEle: Given Ay is empty"
AyLasEle = Ay(UB(Ay))
End Function

Function AyLik(Ay, Lik$)
Dim O$()
AyLik = AySel(Ay, "StrLik", Lik)
End Function

Sub AyLik__Tst()
End Sub

Function AyMak(Ay_or_Itm)
If IsArray(Ay_or_Itm) Then AyMak = Ay_or_Itm: Exit Function

If IsMissing(Ay_or_Itm) Then
    Dim O()
    AyMak = O
    Exit Function
End If
AyMak = Array(Ay_or_Itm)
End Function

Function AyMap(Ay, Fn$, ParamArray Ap())
Dim Av(): Av = Ap
Dim A()
AyMap = AyMapInto_PrmAv(Ay, A, Fn, Av)
End Function

Function AyMapInto(Ay, Into, Fn$, ParamArray Ap())
Dim Av(): Av = Ap
AyMapInto = AyMapInto_PrmAv(Ay, Into, Fn, Av)
End Function

Function AyMapInto_PrmAv(Ay, Into, Fn$, Av())
Erase Into
Dim U&: U = UB(Ay): If U = -1 Then AyMapInto_PrmAv = Into: Exit Function
ReDim Into(U)
Av = AyInsAt(Av, 0, Empty)
Dim J&, M
For J = 0 To U
    If IsObject(Ay(J)) Then
        Set Av(0) = Ay(J)
    Else
        Av(0) = Ay(J)
    End If
    M = RunAv(Fn, Av)
    Into(J) = M
Next
AyMapInto_PrmAv = Into
End Function

Function AyMapIntoSy(Ay, Fn$, ParamArray Ap())
Dim Av(): Av = Ap
Dim Sy$()
AyMapIntoSy = AyMapInto_PrmAv(Ay, Sy, Fn, Av)
End Function

Function AyMax(Ay)
Dim U&: U = UB(Ay): If U = -1 Then Er "Given Ay is Empty, cannot find AyMax"
Dim O: O = Ay(0)
Dim J&
For J = 1 To U
    If Ay(J) > O Then O = Ay(J)
Next
AyMax = O
End Function

Function AyMin(Ay)
Dim U&: U = UB(Ay): If U = -1 Then Er "Given Ay is Empty, cannot find AyMin"
Dim O: O = Ay(0)
Dim J&
For J = 1 To U
    If Ay(J) < O Then O = Ay(J)
Next
AyMin = O
End Function

Function AyMinus(Ay, Ay1, ParamArray Ap())
Dim O: O = AyMinus1(Ay, Ay1)
Dim Av(): Av = Ap
Dim J%
For J = 0 To UB(Av)
    O = AyMinus1(O, Av(J))
    If AyIsEmpty(O) Then Exit For
Next
AyMinus = O
End Function

Function AyMinus__Tst()
Dim mAy1$(10), J%
For J = 0 To 10
    mAy1(J) = J
Next
Dim mAy2$()
Dim mAy$()
mAy = AyMinus(mAy, mAy1, mAy2)
Debug.Print "mAy1=" & Join(mAy1, CtComma)
Debug.Print "mAy=" & Join(mAy, CtComma)
mAy1(0) = "xx"
Debug.Print "mAy1(0)=" & mAy1(0)
Debug.Print "mAy(0)=" & mAy(0)
End Function

Function AyMinus1(Ay1, Ay2)
If AyIsEmpty(Ay1) Then AyMinus1 = Ay1: Exit Function
If AyIsEmpty(Ay2) Then AyMinus1 = Ay1: Exit Function
Dim O: O = Ay1: Erase O
Dim I
For Each I In Ay1
    If Not AyHas(Ay2, I) Then
        Push O, I
    End If
Next
AyMinus1 = O
End Function

Function AyQuote(Ay, QStr$) As String()
Dim O$()
AyQuote = AyMapInto(Ay, O, "Quote", QStr)
End Function

Function AyRev(Ay)
Dim O: O = Ay
Dim J&, U&
U = UB(Ay)
For J = 0 To U
    O(U - J) = Ay(J)
Next
AyRev = O
End Function

Function AyRmvAt(Ay, Optional At&, Optional N& = 1)
Dim U&: U = UB(Ay)
If 0 > At Or At > U Then AyRmvAt = Ay: Exit Function
Dim O: O = Ay: Erase O
Dim J&
For J = 0 To At - 1
    Push O, Ay(J)
Next
For J = At + N To U
    Push O, Ay(J)
Next
AyRmvAt = O
End Function

Function AyRmvBlank(Ay)
Dim O: O = Ay: Erase O
If AyIsEmpty(Ay) Then AyRmvBlank = O: Exit Function
Dim I
For Each I In Ay
    If Not VarIsBlank(I) Then Push O, I
Next
AyRmvBlank = O
End Function

Function AyRmvDup(Ay)
Dim O: O = Ay
Dim J&
For J = 0 To UB(Ay)
    If Not AyHas(O, Ay(J)) Then Push O, Ay(J)
Next
AyRmvDup = O
End Function

Function AyRmvEle(Ay, Ele)
Dim O: O = Ay: Erase O
If Not AyIsEmpty(Ay) Then
    Dim I
    For Each I In Ay
        If I <> Ele Then Push O, I
    Next
End If
AyRmvEle = O
End Function

Function AyRmvFstChr(Ay) As String()
AyRmvFstChr = AyMapIntoSy(Ay, "RmvFstChr")
End Function

Function AyRmvFstEle(Ay)
AyRmvFstEle = AyRmvAt(Ay)
End Function

Function AyRmvLasEle(Ay)
ReSz Ay, UB(Ay) - 1
AyRmvLasEle = Ay
End Function

Function AySel(Ay, Fct$, ParamArray Ap())
If AyIsEmpty(Ay) Then AySel = Ay: Exit Function
Dim O: O = Ay: Erase O
Dim I
Dim Av(): Av = Ap: Av = AyInsAt(Av)
For Each I In Ay
    VarAsg I, Av(0)
    If RunAv(Fct, Av) Then
        Push O, I
    End If
Next
AySel = O
End Function

Function AySel_Idx(Ay, IdxAy&())
If AyIsEmpty(Ay) Then AySel_Idx = Ay: Exit Function
Dim O: O = Ay: Erase O
If AyIsEmpty(IdxAy) Then AySel_Idx = O: Exit Function
Dim I
For Each I In IdxAy
    Push O, Ay(I)
Next
AySel_Idx = O
End Function

Sub AySet(OAy, Idx&, V)
If 0 > Idx Then Er "AySet: {Idx} cannot be -ve", Idx
Dim U&: U = UB(OAy)
If Idx > U Then ReDim Preserve OAy(Idx)
If IsObject(V) Then
    Set OAy(Idx) = V
Else
    OAy(Idx) = V
End If
End Sub

Sub AySet__Tst()
'1 Declare
Dim Ay
Dim Idx&
Dim V

'2 Assign
Ay = Array(1, 2, 3)
Idx = 1
V = "x"

'3 Calling
AySet Ay, Idx, V

'4 Assert
Debug.Assert Sz(Ay) = 3
Debug.Assert Ay(0) = 1
Debug.Assert Ay(1) = "x"
Debug.Assert Ay(2) = 3
End Sub

Function AyShift(OAy)
If AyIsEmpty(OAy) Then Er "AyShift: Given Ay is empty"
AyShift = OAy(0)
OAy = AyRmvAt(OAy)
End Function

Sub AyShift__Tst()
Dim Ay: Ay = Array(1, 2, 3, 4)
Debug.Assert AyShift(Ay) = 1
AyAsstEqExa Ay, Array(2, 3, 4)
End Sub

Function AySlice(Ay, FmIdx&, Optional ToIdx& = -1)
Dim T&
    If ToIdx = -1 Then
        T = UB(Ay)
    Else
        T = ToIdx
    End If
    
Dim U&: U = T - FmIdx
AySlice = AySliceFmU(Ay, FmIdx, U)
End Function

Function AySliceFmU(Ay, Fm&, U&)
Dim O: O = Ay: ReDim O(U)
Dim I&
For I = 0 To U
    O(I) = Ay(I + Fm)
Next
AySliceFmU = O
End Function

Function AySq(Ay, NCol&, Optional NRow&)
'Aim: Join {pLoLin} into at most {pMaxLin} lines in column format
'     Eg. There are 25 lines: Line NN ---.  To join these line by using pMaxLin=10
'         Gives
'                Line 01 ---         Line 11 ---         Line 21 ---
'                Line 02 ---         Line 12 ---         Line 22 ---
'                Line 03 ---         Line 13 ---         Line 23 ---
'                Line 04 ---         Line 14 ---         Line 24 ---
'                Line 05 ---         Line 15 ---         Line 25 ---
'                Line 06 ---         Line 16 ---
'                Line 07 ---         Line 17 ---
'                Line 08 ---         Line 18 ---
'                Line 09 ---         Line 19 ---
'                Line 10 ---         Line 20 ---
End Function

Function AySrt(Ay, Optional IsDes As Boolean)
Dim UR&: UR = UB(Ay): If UR <= 0 Then AySrt = Ay: Exit Function
Dim I&(): I = AySrtIdx(Ay, IsDes)
Dim O: O = Ay: ReDim O(UR)
Dim R&
For R = 0 To UR
    O(R) = Ay(I(R))
Next
AySrt = O
End Function

Function AySrtIdx(Ay, Optional IsDes As Boolean) As Long()
Dim UR&: UR = UB(Ay): If UR < 0 Then Exit Function
Dim O&()
Dim R&, Idx&, J&, V
If IsDes Then
    For R = 0 To UR
        V = Ay(R)
        For J = 0 To UB(O)
            If V > Ay(O(J)) Then O = AyInsAt(O, J, R): GoTo Nxt1
        Next
        Push O, R
Nxt1:
    Next
Else
    For R = 0 To UR
        V = Ay(R)
        For J = 0 To UB(O)
            If Ay(O(J)) > V Then O = AyInsAt(O, J, R): GoTo Nxt2
        Next
        Push O, R
Nxt2:
    Next
End If
AySrtIdx = O
End Function

Function AyStrBrk(Ay, Optional BrkChr$ = ".") As Variant()
Dim O(), U&, J&
For J = 0 To UB(Ay)
    With StrBrk(Ay(J), BrkChr)
        Push O, Array(.S1, .S2)
    End With
Next
AyStrBrk = O
End Function

Function AySubsetIdxAy(Ay, SubAy) As Long()
Dim U&: U = UB(SubAy)
Dim O&(): ReSz O, U
Dim J&
For J = 0 To U
    O(J) = AyIdx(Ay, SubAy(J))
Next
AySubsetIdxAy = O
End Function

Function AySy(Ay) As String()
Dim U&: U = UB(Ay)
Dim O$(): ReSz O, U
Dim J&
For J = 0 To U
    O(J) = VarToStr(Ay(J))
Next
AySy = O
End Function

Function AyTakMaxEle(Ay1, Ay2)
Dim O: O = Ay1
Dim U1&, U2&, U&
    U1 = UB(Ay1)
    U2 = UB(Ay2)
U = Max(U1, U2)
If U1 <> U2 Then ReDim Preserve O(U)
Dim J&
For J = 0 To U
    If J > U1 Then
        O(J) = Ay2(J)
    ElseIf J > U2 Then
        O(J) = Ay1(J)
    Else
        If Ay2(J) > Ay1(J) Then O(J) = Ay2(J)
    End If
Next
AyTakMaxEle = O
End Function

Function AyTrim(Ay) As String()
Dim U&: U = UB(Ay)
Dim O$(): ReSz O, U
Dim J&
For J = 0 To U
    O(J) = Trim(Ay(J))
Next
AyTrim = O
End Function

Function AyUnion(Ay1, Ay2)
Dim O, I, U&, J&
O = Ay1
U = UB(Ay2)
If U >= 0 Then
    For J = 0 To UB(Ay2)
        PushNoDup O, Ay2(J)
    Next
End If
AyUnion = O
End Function

Sub AyWrt(Ay, Ft)
Dim F%: F = FtOpnOup(Ft)
AyWrtFno Ay, F
Close #F
End Sub

Sub AyWrtFno(Ay, Fno%)
Dim J&
For J = 0 To UB(Ay)
    Print #Fno, Ay(J)
Next
End Sub

Function LasEle(Ay)
LasEle = Ay(UB(Ay))
End Function

Function LasObjEle(Ay)
Set LasObjEle = Ay(UB(Ay))
End Function

Function Pop(Ay)
Pop = LasEle(Ay)
AyRmvLasEle Ay
End Function

Function PopObj(Ay)
Set PopObj = LasObjEle(Ay)
AyRmvLasEle Ay
End Function

Sub Push(Ay, V)
Dim N&: N = Sz(Ay)
ReDim Preserve Ay(N)
If IsObject(V) Then
    Set Ay(N) = V
Else
    Ay(N) = V
End If
End Sub

Sub PushAy(OAy, Ay)
Dim J&
For J = 0 To UB(Ay)
    Push OAy, Ay(J)
Next
End Sub

Sub PushAyNoDup(OAy, Ay)
Dim J&
For J = 0 To UB(Ay)
    PushNoDup OAy, Ay(J)
Next
End Sub

Sub PushNoBlank(Ay, I)
If Not VarIsBlank(I) Then Push Ay, I
End Sub

Sub PushNoDup(Ay, I)
If AyHas(Ay, I) Then Exit Sub
Push Ay, I
End Sub

Sub PushObj(Ay, I)
Dim N&
    N = Sz(Ay)
ReDim Preserve Ay(N)
Set Ay(N) = I
End Sub

Sub ReSz(OAy, U&)
If U = -1 Then
    Erase OAy
    Exit Sub
End If
ReDim Preserve OAy(U)
End Sub

Function SyOpt(A) As String()
If IsMissing(A) Then Exit Function
SyOpt = A
End Function

Function Sz&(Ay)
On Error Resume Next
Sz = UBound(Ay) + 1
End Function

Function UB&(Ay)
UB = Sz(Ay) - 1
End Function

Private Function AyChkEqDifIdx(Ay1, Ay2, IsExt As Boolean) As DifIdx
'Assume Ay1 & 2 are same size
Dim U1&, U2&: U1 = UB(Ay1): U2 = UB(Ay2)
If U1 <> U2 Then Er "Prm Err: Given Ay1 and Ay2 should have eq {U1} and {U2} (AyChkEq_2)", U1, U2
Dim O As DifIdx
Dim J&
For J = 0 To UB(Ay1)
    If Not VarIsEq(Ay1(J), Ay2(J), IsExt) Then
        Push O.DifIdx, J
        If Sz(O.DifIdx) >= 10 Then O.IsMore = True: Exit For
    End If
Next
AyChkEqDifIdx = O
End Function

Private Function AyChkEqEr(Ay1, Ay2, IsExa As Boolean) As Dt
Dim O As Dt
    O = AyChkEqSz(Ay1, Ay2)
If ErIsSom(O) Then AyChkEqEr = O: Exit Function


AyChkEqEr = AyChkEqEr_(Ay1, Ay2, IsExa)
End Function

Private Function AyChkEqEr_(Ay1, Ay2, IsExa As Boolean) As Dt
Dim A As DifIdx:
    A = AyChkEqDifIdx(Ay1, Ay2, IsExa)

'Assume Ay1 & Ay2 are same size
Dim U&: U = UB(A.DifIdx)
If U = -1 Then Exit Function
Dim B$: If IsExa Then B = "exactly "
Dim A1$: A1 = FmtQQ("Given 2 Ay of {U} are not ?equal", B)
Dim O As Dt: O = ErNew(A1, UB(Ay1))
If A.IsMore Then
    O = ErApd(O, ".There are more than 10 elements are different")
Else
    O = ErApd(O, ".There are {N} elements are different", U + 1)
End If
Dim I&, V1, V2, J&, Msg$

For J = 0 To U
    I = A.DifIdx(J)
    V1 = Ay1(I)
    V2 = Ay2(I)
    Msg = FmtQQ(".. Ele-{?} of {V1} {V2} {Ty1} {Ty2} are diff", J)
    O = ErApd(O, Msg, I, V1, V2, VarVbTyStr(V1), VarVbTyStr(V2))            '<===
Next
AyChkEqEr_ = O
End Function
