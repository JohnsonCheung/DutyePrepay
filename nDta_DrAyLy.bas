Attribute VB_Name = "nDta_DrAyLy"
Option Compare Database
Option Explicit

Function DrAyLy(DrAy(), Optional BrkAtColIdx% = -1) As String()
Dim W%(): W = DrAyWdtAy(DrAy)
DrAyLy = DrAyLyByWdtAy(DrAy, W, BrkAtColIdx)
End Function

Sub DrAyLy__Tst()
Dim DrAy(): DrAy = Array(Array(1, 2, 3), Array(2, 3))
AyBrw DrAyLy(DrAy)
End Sub

Function DrAyLyByWdtAy(DrAy(), WdtAy%(), Optional BrkLinColIdx% = -1) As String()
Dim J&, O$(), Dr
Dim L$: L = WdtAyLin(WdtAy)
Push O, L
For J = 0 To UB(DrAy)
    Dr = DrAy(J)
    Push O, DrToStr(Dr, WdtAy)
Next
Push O, L

DrAyLyByWdtAy = DrAyLyInsBrkLin(O, BrkLinColIdx)
End Function

Function DrAyLyInsBrkLin(DrAyLy$(), BrkLinColIdx%) As String()
Dim L1$: L1 = DrAyLy(0)
Dim DashAy$(): DashAy = AyRmvAt(Split(L1, "|"), 0)
Dim Idx%: Idx = BrkLinColIdx
Dim UCol%: UCol = UB(DashAy)
If Not (0 <= Idx And Idx <= UCol) Then DrAyLyInsBrkLin = DrAyLy: Exit Function
Dim O$()
Dim J&, Las$, Cur$
Dim Fm&, N&: AyAsg DashAyFmN(DashAy, Idx), Fm, N
Las = Mid(DrAyLy(1), Fm, N)
Push O, L1
For J = 1 To UB(DrAyLy) - 1
    Cur = Mid(DrAyLy(J), Fm, N)
    If Las <> Cur Then
        Push O, L1
        Las = Cur
    End If
    Push O, DrAyLy(J)
Next
Push O, LasEle(DrAyLy)
DrAyLyInsBrkLin = O
End Function

Private Function DashAyFmN(DashAy$(), Idx%) As Long()
Dim J%, Fm&
Fm = 1
For J = 0 To Idx - 1
    Fm = Fm + Len(DashAy(J)) + 1
Next
DashAyFmN = FmNNew(Fm, Len(DashAy(Idx)))
End Function
