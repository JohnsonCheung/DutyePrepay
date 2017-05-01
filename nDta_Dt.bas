Attribute VB_Name = "nDta_Dt"
Option Compare Database
Option Explicit
Type Dt
    Tn As String
    Fny() As String
    DrAy() As Variant
End Type

Sub AA0()
DtAsstEq__Tst
End Sub

Function DtAddCol_Idx(A As Dt) As Dt
Dim Fny$(), DrAy()
    Fny = AyInsAt(A.Fny, 0, "Idx")
    DrAy = DrAyAddCol_Idx(A.DrAy)
DtAddCol_Idx = DtNew(Fny, DrAy)
End Function

Function DtApdDr(A As Dt, Dr) As Dt
If UB(Dr) > UB(A.Fny) Then Er "DtApdDr: Given Tbl has Fields less than given Dr.  Dr cannot Append to Tbl", UB(A.Fny), UB(Dr)
Dim O As Dt: O = A
Push O.DrAy, Dr
DtApdDr = O
End Function

Sub DtApdDr__Tst()
Dim A As Dt: A = DtNew(LvsSplit("A B"), Array(Array(1, 2)))
Dim B As Dt: B = DtApdDr(A, Array(3, 4))
DtBrw A
DtBrw B
End Sub

Sub DtAsgCol(Dt As Dt, Fld$(), ParamArray OCol())
Dim UF%: UF = UB(Fld)
If UF = -1 Then Er "{Fld} has no element"
Dim IAy&(): IAy = AySubsetIdxAy(Dt.Fny, Fld)
Dim D(): D = Dt.DrAy
Dim Av()
    ReDim Av(UF)
    Dim J%
    For J% = 0 To UF
        Dim I&: I = IAy(J)
        Av(J) = DrAyCol_Into(D, OCol(J), I)
    Next
    
For J = 0 To UF
    OCol(J) = Av(J)
Next
End Sub

Sub DtAsstEq(D1 As Dt, D2 As Dt, ParamArray MsgAp())
Dim Av(): Av = MsgAp
Dim Er(): Er = DtChkEq(D1, D2)
If AyHasEle(Er) Then
    DtBrw D1, "Dt1-of-Cmp-2-Dt"
    DtBrw D2, "Dt2-of-Cmp-2-Dt"
End If
ErAsst Er, Av
End Sub

Sub DtAsstEq__Tst()
'1 Declare
Dim D1 As Dt
Dim D2 As Dt

'Assign & Call
'=======================
D1 = DtNewSclVBar("Tbl;AAA|Fld;A;B;C|;1;2;3|;4;5;6|;1;2;3|;4;5;6|;1;2;3|;4;5;6")
D2 = DtNewSclVBar("Tbl;AAA|Fld;A;B;C|;1;2;3|;4;5;6|;1;2;3|;4;5;6|;1;2;3|;4;5;6")
'GoSub ShouldNotThrowEr '3 Calling
'=======================
D1 = DtNewSclVBar("Tbl;AAA|Fld;A;B;C|;1;2;3|;4;5;6|;1;2;3|;4;5;6|;1;2;3|;4;5;6")
D2 = DtNewSclVBar("Tbl;AAA|Fld;A;B;C|;1;2;3|;4;5;7|;1;2;3|;4;5;6|;1;2;3|;4;5;9")
GoSub ShouldThrowEr '3 Calling
Exit Sub
ShouldNotThrowEr:
    On Error GoTo Er1
    DtAsstEq D1, D2
    Return
Er1: Debug.Assert False
ShouldThrowEr:
    On Error GoTo Er2
    DtAsstEq D1, D2, "DtAsstEq__Tst: should ignore"
    Debug.Assert False
    Return
Er2:
End Sub

Sub DtBrw(Dt As Dt, Optional TmpFilPfx$ = "Dt", Optional NoIdx As Boolean, Optional BrkLinFld$)
If Not NoIdx Then Dt = DtAddCol_Idx(Dt)
Dim Ly$(): Ly = DtLy(Dt, BrkLinFld)
AyBrw Ly, TmpFilPfx
End Sub

Sub DtBrw__Tst()
DtBrw DtNew(ApSy("a", "b", "d"), Array(Array(1, 2, 3), Array("asdf", "fdfd", "dfdfdd")))
End Sub

Sub DtBrw1(Dt As Dt, Optional TmpFilPfx$ = "Dt", Optional NoIdx As Boolean, Optional BrkLinFld$, Optional TmpSubFdr$)
HtmBrw DtHtm(Dt, NoIdx, BrkLinFld), TmpFilPfx, TmpSubFdr
End Sub

Function DtChkEq(D1 As Dt, D2 As Dt, ParamArray MsgAp()) As Variant()
Dim Er()
'===============
Dim A_Er_DifFny()
    Er = AyChkSam(D1.Fny, D2.Fny)
    If AyHasEle(Er) Then
        A_Er_DifFny = AyAdd(ErNew("Fields are diff:"), Er)
    End If

Dim A_Er_DifNRow()
    If DtNRec(D1) <> DtNRec(D2) Then
        A_Er_DifNRow = ErNew("NRec are diff {N1} and {N2}", D1.Tn, DtNRec(D1), DtNRec(D2))
    End If

Dim A_Er_DifRow()
    If AyIsEmpty(A_Er_DifFny) Then
        '-------
        Dim D2DrAy()
            D2DrAy = DtSel(D2, FnyToStr(D1.Fny)).DrAy

        Er = DrAyChkEq(D1.DrAy, D2DrAy)
        If AyHasEle(Er) Then
            A_Er_DifRow = AyAdd(ErNew("Some rows are diff:"), Er)
        End If
    End If
'===============
Dim O()
    O = AyAdd(A_Er_DifFny, A_Er_DifNRow)
    O = AyAdd(O, A_Er_DifRow)
    
If AyHasEle(O) Then
    Er = ErNew("Given 2 Dt are different {Tn1} {Tn2}:", D1.Tn, D2.Tn)
    Er = ErApd(Er, "{NRec1} {Fny1}", DtNRec(D1), FnyToStr(D1.Fny))
    Er = ErApd(Er, "{NRec2} {Fny2}", DtNRec(D1), FnyToStr(D2.Fny))
    O = AyAdd(Er, O)
    Dim Av(): Av = MsgAp
    O = AyAddItm(O, Av)
End If
DtChkEq = O
End Function

Function DtCol(Dt As Dt, Fld$) As Variant()
Dim I&: I = AyIdx(Dt.Fny, Fld)
DtCol = DrAyCol(Dt.DrAy, I)
End Function

Function DtCol_Bool(Dt As Dt, Fld$) As Boolean()
Dim I&: I = AyIdx(Dt.Fny, Fld)
DtCol_Bool = DrAyCol_Bool(Dt.DrAy, I)
End Function

Function DtCol_Dte(Dt As Dt, Fld$) As Date()
Dim I&: I = AyIdx(Dt.Fny, Fld)
DtCol_Dte = DrAyCol_Dte(Dt.DrAy, I)
End Function

Function DtCol_Int(Dt As Dt, Fld$) As Integer()
Dim I&: I = AyIdx(Dt.Fny, Fld)
DtCol_Int = DrAyCol_Int(Dt.DrAy, I)
End Function

Function DtCol_Lng(Dt As Dt, Fld$) As Long()
Dim I&: I = AyIdx(Dt.Fny, Fld)
DtCol_Lng = DrAyCol_Lng(Dt.DrAy, I)
End Function

Function DtCol_Str(Dt As Dt, Fld$) As String()
Dim I&: I = AyIdx(Dt.Fny, Fld)
DtCol_Str = DrAyCol_Str(Dt.DrAy, I)
End Function

Function DtColUB&(A As Dt)
DtColUB = Max(UB(A.Fny), DrAyColUB(A.DrAy))
End Function

Function DtFldCnt&(Dt As Dt)
DtFldCnt = Sz(Dt.Fny)
End Function

Function DtFldTyAy(D As Dt) As DAO.DataTypeEnum()
If DtIsNoRec(D) Then Exit Function
Dim OTy() As DAO.DataTypeEnum
Dim Dr, NoMore As Boolean
For Each Dr In D.DrAy
    OTy = TyAyNew(Dr, OTy)
    If TyAyIsAllTxt(OTy) Then Exit For
Next
DtFldTyAy = OTy
End Function

Function DtFtLy(A As Dt, Optional Tn$ = "Table") As String()
Dim O$()
    Push O, "Tbl;" & Tn
    Push O, Join(A.Fny, ";")
    Dim DrAy(): DrAy = A.DrAy
    Dim J%, Dr
    For J = 0 To UB(DrAy)
        Dr = DrAy(J)
        Push O, DrScl(Dr)
    Next
DtFtLy = O
End Function

Function DtHasNoRec(A As Dt) As Boolean
DtHasNoRec = DtNRec(A) = 0
End Function

Function DtHtm$(Dt As Dt, Optional NoIdx As Boolean, Optional BrkLinFld$)
If Not NoIdx Then Dt = DtAddCol_Idx(Dt)
Dim O$():
Push O, "<html><table>"
Push O, FnyHtm(Dt.Fny)
Push O, DrAyHtm(Dt.DrAy)
Push O, "</table><html>"
DtHtm = LyJn(O)
End Function

Function DtInsDr(A As Dt, Dr, At&) As Dt
Dim O As Dt: O = A
O.DrAy = AyInsAt(A.DrAy, At, Dr)
DtInsDr = O
End Function

Function DtIsNoRec(A As Dt) As Boolean
DtIsNoRec = Sz(A.DrAy) = 0
End Function

Function DtLy(A As Dt, Optional BrkLinFldNm$) As String()
Dim Fny$()
Dim DrAy()
    Fny = A.Fny:
    DrAy = DrAyStrCell(A.DrAy)
Dim W%():
    Dim W1%():    W1 = DrAyWdtAy(DrAy)
    Dim W2%():    W2 = AyMapInto(Fny, W2, "StrLen")
    W = AyTakMaxEle(W1, W2)

Dim L$
Dim H$
Dim R$()
    L = WdtAyLin(W)
    H = WdtAyHdr(W, Fny)
    Dim C%:
    C = AyIdx(A.Fny, BrkLinFldNm)
    R = DrAyLyByWdtAy(DrAy, W, C)
DtLy = ApSy(L, H, R)
End Function

Sub DtLy__Tst()
Dim Dt As Dt
Dt = DtNew(ApSy("Msg", "V0"), Array(Array("{TthNm_Sfx} does have Sfx-[_Tst]", "lsdf_Tst")))
AyBrw DtLy(Dt)
End Sub

Function DtNew(Fny$(), DrAy, Optional Tn$ = "Table") As Dt
Dim O As Dt
O.Tn = Tn
O.Fny = Fny
O.DrAy = DrAy
DtNew = O
End Function

Function DtNewSq(DtSq, Optional Tn$ = "Table") As Dt
Dim OFny$()
    OFny = AySy(SqDr(DtSq, 1))
Dim ODrAy()
    ODrAy = SqDrAy_FmTo(DtSq, 2, UBound(DtSq, 1))
DtNewSq = DtNew(OFny, ODrAy, Tn)
End Function

Sub DtNewSq__Tst()
Dim D1 As Dt
Dim S
Dim D2 As Dt
    D1 = DtSample1
    S = DtSq(D1)
    D2 = DtNewSq(S, D1.Tn)
DtAsstEq D1, D2
End Sub

Function DtNRec&(A As Dt)
DtNRec = Sz(A.DrAy)
End Function

Sub DtPush(OAy() As Dt, Dt As Dt)
Dim N&: N = DtSz(OAy)
ReDim Preserve OAy(N)
OAy(N) = Dt
End Sub

Function DtSample1() As Dt
DtSample1 = DtNew(LvsSplit("a b c d"), Array(Array(1, 2, 3, 4), Array(11, 12, 13, 14)))
End Function

Function DtSample2() As Dt
DtSample2 = DtNew(LvsSplit("a b c d"), Array(Array(1, 2, 3, 4), Array(11, 12, 13, 14), , Array(21, 22, 23, 24)))
End Function

Function DtSel(A As Dt, StarFnStr$, Optional Tn$ = "Sel") As Dt
Dim OFny$()
    OFny = FnySel(A.Fny, StarFnStr)
If AyIsEq(A.Fny, OFny) Then
    DtSel = A
    Exit Function
End If
'==========================
Dim I&()
    I = AySubsetIdxAy(A.Fny, OFny)
    Dim E()
    E = AyChkZerOrPos(I)
    E = ErApd(E, "DtSel: Given {StarFnStr} has fields not in Dt-{Flds}", StarFnStr, FnyToStr(A.Fny))
    ErAsst E
'==========================
Dim ODrAy()
    Dim UR&
    Dim R&, Dr
    UR = UB(A.DrAy)
    ReSz ODrAy, UR
    For R = 0 To UR
        Dr = A.DrAy(R)
        ODrAy(R) = DrSel(Dr, I)
    Next
DtSel = DtNew(OFny, ODrAy, Tn)
End Function

Sub DtSel__Tst()
DtBrw DtSrt(DtSel(MthDt_Md(MdCur), "Mdy Nm Ty *"), "Mdy Nm Ty")
DtBrw DtSrt(DtSel(MthDt_Md(MdCur), "Mdy Nm Ty"), "Mdy Nm Ty")
DtBrw DtSrt(DtSel(MthDt_Md(MdCur), "Ty Mdy Nm Ty asdf"), "Mdy Nm Ty")
End Sub

Function DtSrt(Dt As Dt, FnStr$) As Dt
Dim UR&: UR = DtNRec(Dt) - 1: If UR = 0 Then DtSrt = Dt: Exit Function
Dim F$(): F = NmBrk(FnStr)
Dim UF%: UF = UB(F)
Dim IsDesAy() As Boolean
    Dim J%
    ReDim IsDesAy(UF)
    For J = 0 To UB(F)
        If LasChr(F(J)) = "-" Then
            IsDesAy(J) = True
            F(J) = RmvLasChr(F(J))
        End If
    Next

Dim O(): ReDim O(UR)
    Dim ColIdxAy&(): ColIdxAy = AySubsetIdxAy(Dt.Fny, F)
    Dim I&(): I = DrAySrtIdx(Dt.DrAy, ColIdxAy, IsDesAy)
    Dim D(): D = Dt.DrAy
    Dim R&
    For R = 0 To UR
        O(R) = D(I(R))
    Next
DtSrt = DtNew(Dt.Fny, O)
End Function

Sub DtSrt__Tst()
Dim A As Dt: A = DtSample2
DtBrw DtSrt(A, "a")
End Sub

Function DtSz&(Ay() As Dt)
On Error Resume Next
DtSz = UBound(Ay) + 1
End Function

Function DtToStr$(A As Dt, Optional Tn$ = "Table")
DtToStr = LyJn(DtLy(A, Tn))
End Function

Sub DtToStr__Tst()
StrBrw DtToStr(DtSample1)
End Sub

Function DtUB&(Ay() As Dt)
DtUB = DtSz(Ay) - 1
End Function

Function DtUnion(D1 As Dt, D2 As Dt, Optional Tn$ = "Union") As Dt
Dim A_D2DrAy()
    A_D2DrAy = D2.DrAy ' The D2 DrAy to append to end OFny O.DrAy, which has been assigned by D1.DrAy

Dim A_Fny$()
    Dim F2$(): F2 = D2.Fny
    A_Fny = AyUnion(D1.Fny, F2) '<== OFnyny is set

Dim A_Idx&()
    A_Idx = AyIdxAy(A_Fny, F2)

Dim OFny$()
    OFny = A_Fny
Dim ODrAy():
    ODrAy = D1.DrAy         ' Assign D1.DrAy to ODrAyrAy
    
    Dim D2Dr()                     ' The D1 Dr when looping D1.DrAy
    Dim R&, C, U&
    
    For R = 0 To UB(A_D2DrAy)
        D2Dr = A_D2DrAy(R)
        If Not AyIsEmpty(D2Dr) Then
            U = UB(D2Dr)
            ReDim Dr(U)
            For C = 0 To U
                Dr(A_Idx(C)) = D2Dr(C)
            Next
            Push ODrAy, Dr
        End If
    Next
DtUnion = DtNew(OFny, ODrAy, Tn)
End Function

Sub DtUnion1__Tst()
Dim D1 As Dt: D1 = DtNew(LvsSplit("a b c d"), Array(Array(1, 2, 3, 4), Array("a", "b", "d", "e")))
Dim D2 As Dt: D2 = DtNew(LvsSplit("a b c e"), Array(Array(1, 21, 3, 4), Array("a", "b", "d", "f")))
Dim Exp As Dt
Dim Act As Dt
Act = DtUnion(D1, D2)
If True Then
    DtBrw Act
Else
    Dim Er()
    Er = DtChkEq(Act, Exp)
    ErBrw Er
End If
End Sub

Sub DtUnion2__Tst()
Dim Av()
Dim D1 As Dt: D1 = DtNew(LvsSplit("a b c d"), Av)
Dim D2 As Dt: D2 = DtNew(LvsSplit("a b c e"), Array(Array(1, 21, 3, 1), Array("a", "b", "d")))
Dim Exp As Dt
Dim Act As Dt
Act = DtUnion(D1, D2)
If True Then
    DtBrw Act
Else
    ErBrw DtChkEq(Act, Exp)
End If
End Sub

Function DtURec&(A As Dt)
DtURec = DtNRec(A) - 1
End Function

Function DtWhere(D As Dt, Fld$, Cndn$, Optional Ty As VbVarType = VbVarType.vbInteger) As Dt
Dim DrAy(): DrAy = D.DrAy
Dim OD(), Dr, V
Dim Idx%: Idx = AyIdx(D.Fny, Fld)
Dim R&

Dim A$
Select Case Ty
Case VbVarType.vbInteger: A = "?" & Cndn
Case VbVarType.vbString: A = """?""" & Cndn
Case VbVarType.vbDate: A = "#?#" & Cndn
Case Else
    Stop
End Select

For R = 0 To UB(D.DrAy)
    Dr = DrAy(R)
    V = Dr(Idx)
    If Eval(FmtQQ(A, V)) Then Push OD, Dr
Next
Dim O As Dt
O.Fny = D.Fny
O.DrAy = OD
DtWhere = O
End Function

Sub DtWhere__Tst()
Dim D As Dt
D = DtNew(LvsSplit("a b c"), Array(Array(1, 2, 3), Array(2, 3, 4)))
'DtBrw DtWhere(D, "a", "=1")

D = DtNew(LvsSplit("a b c"), Array(Array("1", 2, 3), Array("2", 3, 4)))
'DtBrw DtWhere(D, "a", "=""1""", vbString)

D = DtNew(LvsSplit("a b c"), Array(Array(#3/2/2017#, 2, 3), Array(#3/2/2017#, 3, 4)))
DtBrw DtWhere(D, "a", "=#3/2/2017#", vbDate)
End Sub

Private Function TyAyIsAllTxt(Ty() As DataTypeEnum) As Boolean
Dim J&
For J = 0 To UB(Ty)
    If Ty(J) <> dbText Then Exit Function
Next
TyAyIsAllTxt = True
End Function

Private Function TyAyNew(Dr, OldTyAy() As DataTypeEnum) As DataTypeEnum()
Dim U1&, U2&: U1 = UB(OldTyAy): U2 = UB(Dr)
Dim O() As DataTypeEnum: O = OldTyAy
If U2 > U1 Then ReDim Preserve O(U2)
Dim J&
For J = 0 To U2
    If O(J) <> VarDaoTy(Dr(J)) Then
        O(J) = dbText
    End If
Next
TyAyNew = O
End Function

Private Sub TyAyNew__Tst()
Dim A() As DataTypeEnum
Dim Dr(): Dr = Array(1, 2, 3, "4", CLng(5))
Dim B() As DataTypeEnum
B = TyAyNew(Dr, A)
Stop
End Sub
