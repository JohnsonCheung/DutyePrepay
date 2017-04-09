Attribute VB_Name = "nXls_Sq"
Option Compare Database
Option Explicit

Sub SqBrw(Sq, Optional Pfx$ = "Sq")
If SqIsEmpty(Sq) Then
    If MsgBox("Given Sq of Pfx-[" & Sq & "] is empty.  Stop?", vbQuestion + vbYesNo) = vbYes Then Stop
    Exit Sub
End If
DrAyBrw SqDrAy(Sq), , Pfx
End Sub

Function SqDr(Sq, R)
Dim O: O = Sq: Erase O
ReDim O(UBound(Sq, 2) - 1)
Dim J&
For J = 0 To UBound(O)
    O(J) = Sq(R, J + 1)
Next
SqDr = O
End Function

Function SqDrAy(Sq) As Variant()
Dim O()
Dim U&: U = UBound(Sq, 1) - 1
ReSz O, U
Dim J&
For J = 0 To U
    O(J) = SqDr(Sq, J + 1)
Next
SqDrAy = O
End Function

Function SqIsEmpty(Sq) As Boolean
If IsEmpty(Sq) Then SqIsEmpty = True: Exit Function
VarAsstSq Sq
On Error GoTo X
If UBound(Sq, 2) <= 0 Then Exit Function
If UBound(Sq, 1) <= 0 Then Exit Function
Exit Function
X: SqIsEmpty = True
End Function

Sub SqIsEmpty__Tst()
Dim Sq()
Debug.Assert SqIsEmpty(Sq) = True
ReDim Sq(1 To 1, 1 To 1)
Debug.Assert SqIsEmpty(Sq) = False
End Sub

Sub SqPutCell(Sq, Cell As Range)
If SqIsEmpty(Sq) Then Exit Sub
CellReSz(Cell, Sq).Value = Sq
End Sub

Function SqRIdx&(Sq, V, CIdx&)
Dim R&
For R = 1 To UBound(Sq, 1)
    If Sq(R, CIdx) = V Then SqRIdx = R: Exit Function
Next
End Function
