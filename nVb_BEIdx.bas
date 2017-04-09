Attribute VB_Name = "nVb_BEIdx"
Option Compare Database
Option Explicit

Sub BEIdxAsg(BEIdx&(), OB&, OE&)
OB = BEIdx(0)
OE = BEIdx(1)
End Sub

Sub BEIdxDmp(BEIdx&())
Debug.Print BEIdxToStr(BEIdx)
End Sub

Function BEIdxFmN(BEIdx&(), Optional FmIsBase1 As Boolean) As Long()
Dim OffSet%
    If FmIsBase1 Then OffSet = 1 Else OffSet = 0
Dim B&, E&
B = BEIdx(0)
E = BEIdx(1)
BEIdxFmN = ApLngAy(B + OffSet, E - B + 1)
End Function

Function BEIdxNew(B&, E&) As Long()
If B > E Then Er "BEIdxNew: {B} > {E}", B, E
If B < -1 Then Er "BEIDxNew: {B} <-1", B
If E < -1 Then Er "BEIDxNew: {E} <-1", E
BEIdxNew = ApLngAy(B, E)
End Function

Function BEIdxToStr$(BEIdx&())
BEIdxToStr = BEIdx(0) & " " & BEIdx(1)
End Function
