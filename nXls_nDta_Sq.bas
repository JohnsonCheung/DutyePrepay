Attribute VB_Name = "nXls_nDta_Sq"
Option Compare Database
Option Explicit

Sub SqAsstNx2(Sq, ParamArray MsgAp())
Dim Av(): Av = MsgAp
ErAsst SqChkNx2(Sq), Av
End Sub

Sub SqBrw(Sq, Optional Pfx$ = "Sq")
If SqIsEmpty(Sq) Then
    If MsgBox("Given Sq of Pfx-[" & Sq & "] is empty.  Stop?", vbQuestion + vbYesNo) = vbYes Then Stop
    Exit Sub
End If
DrAyBrw SqDrAy(Sq), , Pfx
End Sub

Function SqChkNx2(Sq) As Variant()
On Error GoTo X
If UBound(Sq, 2) = 2 Then Exit Function
SqChkNx2 = ErNew("Given Sq is not Nx2")
Exit Function
X:
SqChkNx2 = ErNew("Given Sq is not Nx2.  It has {error} in accessing UBound(Sq,2).", Err.Description)
End Function

Function SqDic(SqOfNx2) As Dictionary
SqAsstNx2 SqOfNx2
End Function

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

Function SqDrAy_FmTo(Sq, FmIdx&, ToIdx&) As Variant()
Dim O()
Dim U&: U = UBound(Sq, 1) - 2
ReSz O, U
Dim J&
For J = FmIdx To ToIdx
    O(J - FmIdx) = SqDr(Sq, J)
Next
SqDrAy_FmTo = O
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
ReDim Sq(1 To 0, 1 To 1)
End Sub

Function SqPutCell(Sq, Cell As Range, Optional CrtLo As Boolean) As Range
If SqIsEmpty(Sq) Then Exit Function
Dim O As Range: Set O = CellReSz(Cell, Sq)
O.Value = Sq
If CrtLo Then RgLo O
Set SqPutCell = O
End Function

Function SqRIdx&(Sq, V, CIdx&)
Dim R&
For R = 1 To UBound(Sq, 1)
    If Sq(R, CIdx) = V Then SqRIdx = R: Exit Function
Next
End Function

Function SqTranspose(Sq)
Dim NC&, NR&, R&, C&
NC = UBound(Sq, 2)
NR = UBound(Sq, 1)
Dim O()
ReDim O(1 To NC, 1 To NR)
For R = 1 To NR
    For C = 1 To NC
        O(C, R) = Sq(R, C)
    Next
Next
SqTranspose = Sq
End Function
