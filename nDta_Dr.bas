Attribute VB_Name = "nDta_Dr"
Option Compare Database
Option Explicit

Sub DrAsstEq(Dr1, Dr2)
ErAsst DrChkEq(Dr1, Dr2)
End Sub

Function DrChkEq(Dr1, Dr2) As Variant()
Dim S1$, S2$
S1 = DrScl(Dr1)
S2 = DrScl(Dr2)
Dim Er(), Er1()
Er = StrChkEq(S1, S2)
If AyHasEle(Er) Then
    Er1 = ErNew("Two given Dr of {U1} {U2} are different at {FldIdx}:", UB(Dr1), UB(Dr2), DrDifAt(Dr1, Dr2))
    DrChkEq = AyAdd(Er1, Er)
End If
End Function

Sub DrChkEq__Tst()
Dim Dr1: Dr1 = Array(1, 2, 3, 4, 6)
Dim Dr2: Dr2 = Array(1, 2, 3, 4, 5)
DrAsstEq Dr1, Dr2
End Sub

Function DrDifAt&(Dr1, Dr2)
Dim J&, U&
U = Min(UB(Dr1), UB(Dr2))
For J = 0 To U
    If Dr1(J) <> Dr2(J) Then DrDifAt = J: Exit Function
Next
End Function

Function DrIsEmptyRec(Dr()) As Boolean
Dim I
For Each I In Dr
    If Not VarIsBlank(I) Then Exit Function
Next
DrIsEmptyRec = True
End Function

Function DrLasNonBlankIdx&(Dr)
Dim O&
For O = UB(Dr) To 0 Step -1
    If Not VarIsBlank(Dr(O)) Then DrLasNonBlankIdx = O
Next
DrLasNonBlankIdx = -1
End Function

Function DrSel(Dr, Idx&())
Dim U&: U = UB(Idx)
Dim O: O = Dr: ReSz O, U
Dim J&
For J = 0 To U
    O(J) = Dr(Idx(J))
Next
DrSel = O
End Function

Sub DrSel__Tst()
Dim I&(): I = ApLngAy(1, 2, 5)
Dim D(): D = Array("A", "B", "C", "D", "E", "F")
Dim Act(): Act = DrSel(D, I)
Debug.Assert Sz(Act) = 3
Debug.Assert Act(0) = "B"
Debug.Assert Act(1) = "C"
Debug.Assert Act(2) = "F"
End Sub

Function DrToStr$(Dr, WdtAy%())
Dim UW%: UW = UB(WdtAy)
Dim UD%: UD = UB(Dr)
If UD > UW Then
    MsgBox FmtQQ("DrToStr: UD(?) cannot > UW(?)|UD = UB(Dr)|UW = UB(WdtAy)", UD, UW)
    Stop
End If
Dim O$(): ReDim O$(UW)
    Dim J%
    For J = 0 To UD
        O(J) = AlignL(Dr(J), WdtAy(J))
    Next
    For J = UD + 1 To UW
        O(J) = Space(WdtAy(J))
    Next
DrToStr = Quote(Join(O, " | "), "| * |")
End Function
