Attribute VB_Name = "nXls_nDta_Dt"
Option Compare Database
Option Explicit

Sub DtPutCell(A As Dt, Cell As Range)
Dim Sq: Sq = DtSq(A)
CellReSz(Cell, Sq).Value = Sq
End Sub

Function DtSq(A As Dt)
Dim UC&: UC = DtColUB(A)
Dim UR&: UR = UB(A.DrAy)
Dim IR&, IC&, Dr, DrAy(), R&
Dim O(): ReDim O(1 To UR + 2, 1 To UC + 1)
DrAy = A.DrAy
For IC = 0 To UB(A.Fny)
    O(1, IC + 1) = A.Fny(IC)
Next
For IR = 0 To UR
    Dr = DrAy(IR)
    R = IR + 2
    For IC = 0 To UB(Dr)
        O(R, IC + 1) = Dr(IC)
    Next
Next
DtSq = O
End Function

Function DtWs(Dt As Dt, Optional WsNm$ = "Data") As Worksheet
Dim O As Worksheet
Set O = WsNew(WsNm)
DtPutCell Dt, WsA1(O)
Set DtWs = O
End Function
