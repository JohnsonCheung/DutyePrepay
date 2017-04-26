Attribute VB_Name = "nDta_DrAySrt"
Option Compare Database
Option Explicit

Function DrAySrt(DrAy(), ColIdxAy%()) As Variant()
Dim UR&: UR = UB(DrAy): If UR = -1 Then Exit Function
Dim I&(): I = DrAySrtIdx(DrAy, ColIdxAy)
Dim O(): ReDim O(UR)
Dim R&
For R = 0 To UR
    O(R) = DrAy(I(R))
Next
DrAySrt = O
End Function

Function DrAySrtIdx(DrAy(), Optional ByVal ColIdxAy, Optional ByVal IsDesAy) As Long()
Dim UC%:               UC = UB(ColIdxAy):               If UC = -1 Then Exit Function
Dim Des() As Boolean: Des = GetDesAy(IsDesAy, UC)
Dim CIdx&():         CIdx = Prm_CIdx(ColIdxAy)
Dim ColP():          ColP = DrAyCol(DrAy, CIdx(0), ColP) ' First Column of DrAy
Dim O&():               O = AySrtIdx(ColP, Des(0))
Dim GpAyP():        GpAyP = Array(ApLngAy(0, UB(DrAy)))
Dim ColPS():        ColPS = Fnd_Col(ColP, O, GpAyP)      ' P=Prv S=Sorted
Dim ICol%
Dim GpAy()
For ICol = 1 To UC
    GpAy = Fnd_GpAy(ColPS, GpAyP)
           If AyIsEmpty(GpAy) Then DrAySrtIdx = O: Exit Function
    ColP = DrAyCol(DrAy, CIdx(ICol))
   ColPS = Fnd_Col(ColP, O, GpAyP)                       ' Sort {Col} in order of {O} only those elements in {GpAy}
       O = Srt(ColPS, GpAy, Des(ICol), O)
   GpAyP = GpAy
Next
DrAySrtIdx = O
End Function

Sub DrAySrtIdx__Tst()
Dim DrAy(): DrAy = Array( _
    Array(0, 4, "4a"), _
    Array(1, 1, "1c"), _
    Array(2, 1, "1a"), _
    Array(3, 4, "4b"), _
    Array(4, 4, "4c"), _
    Array(5, 4, "4d"), _
    Array(6, 5, "5x"), _
    Array(7, 3, "3a"), _
    Array(8, 1, "1b"), _
    Array(9, 3, "3b"))
Dim Act&(): Act = DrAySrtIdx(DrAy, ApIntAy(1, 2))
Debug.Assert Sz(Act) = 10
Debug.Assert AyHasDup(Act) = False
Debug.Assert AyMax(Act) = 9
Debug.Assert AyMin(Act) = 0
Debug.Assert Act(0) = 2
Debug.Assert Act(1) = 8
Debug.Assert Act(2) = 1
Debug.Assert Act(3) = 7
Debug.Assert Act(4) = 9
Debug.Assert Act(5) = 0
Debug.Assert Act(6) = 3
Debug.Assert Act(7) = 4
Debug.Assert Act(8) = 5
Debug.Assert Act(9) = 6
End Sub

Private Sub CvGpEle_ToFmU(Gp, OFm&, OU&)
Dim oTo&
OFm = Gp(0)
oTo = Gp(1)
OU = oTo - OFm
End Sub

Private Function Fnd_Col(Col(), SortedIdx&(), GpAy()) As Variant()
'Return a {NewCol} from {Col} with those element as pointed by {GpAy} in the order of {SortedIdx}
'Note: UB-of-[Col SortedIdx] are same
Dim U&: U = UB(SortedIdx)
Dim O(): ReSz O, U
Dim R&
Dim J&, I&, Gp&(), Idx&
For J = 0 To UB(GpAy)
    Gp = GpAy(J)
    For I = Gp(0) To Gp(1)
        Idx = SortedIdx(I)
        O(I) = Col(Idx)
    Next
Next
Fnd_Col = O
End Function

Private Function Fnd_GpAy(Col(), GpAy()) As Variant()
Dim O()
Dim Gp
For Each Gp In GpAy
    Dim IFm&, U&: CvGpEle_ToFmU Gp, IFm, U
    Dim C()
        ReDim C(U)
        Dim I&
        For I = 0 To U
            C(I) = Col(I + IFm)
        Next
    Dim OO(): OO = AyGpBySamVal(C, True)
        For I = 0 To UB(OO)
            Gp = OO(I)
            Gp(0) = Gp(0) + IFm
            Gp(1) = Gp(1) + IFm
            OO(I) = Gp
        Next
    PushAy O, OO
Next
Fnd_GpAy = O
End Function

Private Function GetDesAy(IsDesAy, UC%) As Boolean()
Dim O() As Boolean
If IsMissing(IsDesAy) Then
    ReDim O(UC)
Else
    O = IsDesAy
    If UB(O) <> UC Then Er "Given IsDesAy has {UB} should equal {IdxAy-UB}", UB(O), UC
End If
GetDesAy = O
End Function

Private Function Prm_CIdx(ColIdxAy) As Long()
Dim O&()           'ColIdx of each col to be sorted
If IsMissing(ColIdxAy) Then
    ReDim O(0): O(0) = 0
Else
    O = ColIdxAy
End If
Prm_CIdx = O
End Function

Private Function Srt(Col(), GpAy(), IsDes As Boolean, Idx&()) As Long()
Dim O&(): O = Idx
Dim Gp
For Each Gp In GpAy
    Dim IFm&, U&:    CvGpEle_ToFmU Gp, IFm, U
    Dim C():     C = AySliceFmU(Col, IFm, U)
    Dim CI&():  CI = AySrtIdx(C, IsDes)
    Dim CI1&():      ReDim CI1(U)
    Dim I&
    For I = 0 To U
        O(I + IFm) = Idx(CI(I) + IFm)
    Next
Next
Srt = O
End Function

