Attribute VB_Name = "nDta_DrAy"
Option Compare Database
Option Explicit

Function DrAyAddCol_Const(DrAy(), ConstVal) As Variant()
If AyIsEmpty(DrAy) Then Exit Function
Dim O(): O = DrAy
Dim UC&: UC = DrAyColUB(O) + 1
Dim R&
For R = 0 To UB(O)
    ReSz O(R), UC
    O(R)(UC) = ConstVal
Next
DrAyAddCol_Const = O
End Function

Function DrAyAddCol_ConstAtBeg(DrAy(), ConstVal) As Variant()
Dim U&: U = UB(DrAy)
Dim O(): ReSz O, U
Dim R&
For R = 0 To UB(O)
    O(R) = AyInsAt(DrAy(R), , ConstVal)
Next
DrAyAddCol_ConstAtBeg = O
End Function

Function DrAyAddCol_Idx(DrAy()) As Variant()
Dim O(), J&, U&
U = UB(DrAy)
ReSz O, U
For J = 0 To U
    If AyIsEmpty(DrAy(J)) Then
        O(J) = Array(J)
    Else
        O(J) = AyInsAt(DrAy(J), , J)
    End If
Next
DrAyAddCol_Idx = O
End Function

Sub DrAyAsg(DrAy(), ParamArray OAp())
Dim UR&: UR = UB(DrAy)
Dim UC%
    Dim Av()
    Av = OAp
    UC = UB(Av)
Dim I%
    Dim V
    For I = 0 To UC
        V = OAp(I)
        ReDim V(UR)
        OAp(I) = V
    Next
Dim J&, Dr
For J = 0 To UR
    Dr = DrAy(J)
    For I = 0 To UC
        VarAsg Dr(I), OAp(I)(J)
    Next
Next
End Sub

Sub DrAyAsg__Tst()
Dim DrAy()
Dim A%(), B&()
    DrAy = Array(Array(1, 2, 3, 4), Array(5, 6, 7, 8))
    DrAyAsg DrAy, A, B
AyAsstEqExa A, ApIntAy(1, 5), "DrAyAsg__Tst Er"
AyAsstEqExa B, ApLngAy(2, 6), "DrAyAsg__Tst Er"
End Sub

Sub DrAyBrw(DrAy(), Optional NoIdx As Boolean, Optional Fx$ = "DrAy", Optional BrkAtColIdx% = -1)
If Not NoIdx Then DrAy = DrAyAddCol_Idx(DrAy)
AyBrw DrAyLy(DrAy, BrkAtColIdx), Fx
End Sub

Function DrAyChkEq(DrAy1(), DrAy2()) As Variant()
Dim Er() 'shared using
'====
Dim O()
Dim NEr%
    Dim R&, U&, U2&, U1&
    NEr = 0
    U1 = UB(DrAy1)
    U2 = UB(DrAy2)
    If U1 <> U2 Then
        O = ErNew("Given 2 DrAy has different {UB1} and {UB2}", U1, U2)
    End If
    U = Min(U1, U2)
    For R = 0 To U
        Er = DrChkEq(DrAy1(R), DrAy2(R))
        If AyHasEle(Er) Then
            PushAy O, ErNew("**Row-{R} is different", R)
            PushAy O, Er
            NEr = NEr + 1
            If NEr = 10 Then DrAyChkEq = O: Exit For
        End If
    Next

If AyHasEle(O) Then
    Dim A$:
    Dim M$:
    If NEr = 10 Then A = "at least " Else A = ""
    M = FmtQQ("Given 2 DrAy has ?{NEr} rows are different", A)
    Er = ErNew(M, NEr)
    O = AyAdd(Er, O)
End If
DrAyChkEq = O
End Function

Sub DrAyChkEq__Tst()
Dim D1(), D2()
'=========
D1 = Array(Array(1, 2, 3, 4), Array(2, 3, 4, 5))
D2 = Array(Array(1, 2, 3, 4), Array(2, 3, 4, 5))
Debug.Assert AyHasEle(DrAyChkEq(D1, D2)) = False
'=========
D2 = Array(Array(1, 2, 3, 4), Array(2, 3, 4, 6))
ErBrw DrAyChkEq(D1, D2)
End Sub

Function DrAyCol(DrAy(), Optional ColIdx&) As Variant()
DrAyCol = DrAyCol_Into(DrAy, EmptyVarAy, ColIdx)
End Function

Function DrAyCol_Bool(DrAy(), Optional ColIdx&) As Boolean()
DrAyCol_Bool = DrAyCol_Into(DrAy, EmptyBoolAy, ColIdx)
End Function

Function DrAyCol_Dte(DrAy(), Optional ColIdx&) As Date()
DrAyCol_Dte = DrAyCol_Into(DrAy, EmptyDteAy, ColIdx)
End Function

Function DrAyCol_Int(DrAy(), Optional ColIdx&) As Integer()
DrAyCol_Int = DrAyCol_Into(DrAy, EmptyIntAy, ColIdx)
End Function

Function DrAyCol_Into(DrAy(), OIntoAy, Optional ColIdx&)
Dim UR&: UR = UB(DrAy)
ReSz OIntoAy, UR
If Not AyIsEmpty(DrAy) Then
    Dim J&, Dr
    For J = 0 To UR
        Dr = DrAy(J)
        If UB(Dr) >= ColIdx Then
            OIntoAy(J) = DrAy(J)(ColIdx)
        End If
    Next
End If
DrAyCol_Into = OIntoAy
End Function

Function DrAyCol_Lng(DrAy(), Optional ColIdx&) As Long()
DrAyCol_Lng = DrAyCol_Into(DrAy, EmptyLngAy, ColIdx)
End Function

Function DrAyCol_Str(DrAy(), Optional ColIdx&) As String()
DrAyCol_Str = DrAyCol_Into(DrAy, EmptySy, ColIdx)
End Function

Function DrAyColSz&(DrAy())
DrAyColSz = DrAyColUB(DrAy) + 1
End Function

Function DrAyColUB&(DrAy)
Dim O&, I, U&
If AyIsEmpty(DrAy) Then Exit Function
For Each I In DrAy
    O = Max(O, UB(I))
Next
DrAyColUB = O
End Function

Function DrAyHtm$(DrAy())
If AyIsEmpty(DrAy) Then Exit Function
Dim O$(), Dr
For Each Dr In DrAy
    Push O, DrHtm(Dr)
Next
DrAyHtm = LyJn(O)
End Function

Function DrAyNew_AyAp(Ay, ParamArray AyAp()) As Variant()
Dim O(), J&, Dr(), UFld&, I%
Dim Av()
Av = AyAp
UFld = UB(Av)
For J = 0 To UB(Ay)
    Erase Dr
    Push Dr, Ay(J)
    For I = 0 To UFld
        Push Dr, AyAp(I)(J)
    Next
    Push O, Dr
Next
DrAyNew_AyAp = O
End Function

Function DrAyNewKVByFldAp(FnStr$, ParamArray Ap()) As Variant
Dim F$(): 'F = NmstrBrk(FnStr)
Dim J%, O()
For J = 0 To UB(F)
    Push O, Array(F(J), Ap(J))
Next
DrAyNewKVByFldAp = O
End Function

Function DrAyRmvEmptyDr(DrAy()) As Variant()
If AyIsEmpty(DrAy) Then Exit Function
Dim Dr
Dim O()
For Each Dr In DrAy
    If Not AyIsAllEmptyEle(Dr) Then Push O, Dr
Next
DrAyRmvEmptyDr = O
End Function

Function DrAySample() As Variant()
DrAySample = Array(Array(1, 2, 3, 4, 5, 6), Array(2, 3, 4, 5, 6, 7), Array("22", "33", "44", "5"))
End Function

Function DrAySel(DrAy(), IdxAy%())
Dim UR&: UR = UB(DrAy)
Dim UF%: UF = UB(IdxAy)
Dim O(): ReSz O, UR
Dim R&, Dr(), C%, Idx%
For R = 0 To UR
    ReDim ODr(UF)
    Dr = DrAy(R)
    For C = 0 To UF
        Idx = IdxAy(C)
        If Idx <= UF Then
            ODr(C) = Dr(Idx)
        End If
    Next
    O(R) = ODr
Next
DrAySel = O
End Function

Function DrAySq(DrAy())
Dim NC&: NC = DrAyColSz(DrAy): If NC = 0 Then Exit Function
Dim NR&: NR = Sz(DrAy):        If NR = 0 Then Exit Function
Dim O()
ReDim O(1 To NR, 1 To NC)
Dim R&, C&, Dr
For R = 1 To NR
    Dr = DrAy(R - 1)
    For C = 1 To Sz(Dr)
        O(R, C) = Dr(C - 1)
    Next
Next
DrAySq = O
End Function

Sub DrAySq__Tst()
Dim DrAy()
Push DrAy, Array(1, 2, 3, 4)
Push DrAy, Array(2, 4, "a")
Dim Act
Act = DrAySq(DrAy)
Debug.Assert UBound(Act, 1) = 2
Debug.Assert UBound(Act, 2) = 4
Debug.Assert Act(1, 1) = 1
Debug.Assert Act(1, 2) = 2
Debug.Assert Act(1, 3) = 3
Debug.Assert Act(1, 4) = 4

End Sub

Function DrAyStrCell(DrAy()) As Variant()
Dim U&: U = UB(DrAy)
If U = -1 Then Exit Function
Dim O(): ReSz O, U
Dim ODr$()
Dim R&, Dr, UU&, J&
For R = 0 To U
    Dr = DrAy(R)
    UU = UB(Dr)
    If UU >= 0 Then ReDim ODr(UU)
    For J = 0 To UU
        ODr(J) = VarToStr(Dr(J))
    Next
    O(R) = ODr
Next
DrAyStrCell = O
End Function

Function DrAyWdtAy(DrAy()) As Integer()
Dim O%()
Dim U%: U = -1

Dim R&, W%, Dr, U1%, C%

Dim D(): D = DrAy
For R = 0 To UB(D)
    Dr = D(R)
    U1 = UB(Dr)
    If U1 > U Then ReDim Preserve O(U1): U = U1
    For C = 0 To U1
        W = Len(Dr(C))
        If O(C) < W Then O(C) = W
    Next
Next
DrAyWdtAy = O
End Function
