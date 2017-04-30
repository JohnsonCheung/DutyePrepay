Attribute VB_Name = "nXls_nAct_nFmt_Dt"
Option Compare Database
Option Explicit

Sub DtPutCellWithFmt(A As Dt, Cell As Range, Fmtr As LoFmtr)
Dim WithHdr As Boolean
Dim WithLbl As Boolean
    WithHdr = Not AyIsEmpty(Fmtr.HdrSq)
    WithLbl = Not AyIsEmpty(Fmtr.Lbl)
'----------------------------------------------
Dim HdrSq
Dim DtaCell As Range
Dim DtaSq
Dim DtaRg As Range
    'HdrSq
    HdrSq = Fmtr.HdrSq
    'DtaCell
    If WithHdr Then
        Dim HdrNRow%, HdrN, LblN%
        HdrN = UBound(HdrSq, 1)
        LblN = IIf(WithLbl, 1, 0)
        HdrNRow = HdrN + LblN
        Set DtaCell = RgRC(Cell, HdrNRow, 1)
    Else
        Set DtaCell = Cell
    End If
    'DtaSq
    If WithLbl Then
        DtaSq = DtSq(DtNew(Fmtr.Lbl, A.DrAy, A.Tn))
    Else
        DtaSq = DtSq(A)
    End If
    'DtaRg
    Set DtaRg = CellReSz(DtaCell, DtaSq)
'=======================
If WithHdr Then
    SqPutCell HdrSq, Cell
End If
SqPutCell DtaSq, DtaCell
ListObjFmt ListObjCrt(DtaRg), Fmtr
End Sub

Sub DtPutCellWithFmt__Tst()
'1 Declare
Dim A As Dt
Dim Cell As Range
Dim Fmtr As LoFmtr
'-----
Dim Wb As Workbook
    Set Wb = WbNew
'2 Assign ================================================
Dim NFld%
    NFld = 4

A = DtNewSclVBar("Tbl;Tbl|Fld;A;B;C;D;E|;1;2;3;4;5|;A;B;C;D;E|;F;G;H;I;J;K")
Set Cell = WsA1(WbFstWs(Wb))
Fmtr = ZTR_Fmtr

'3 Calling =========
DtPutCellWithFmt A, Cell, Fmtr
Stop
WbClsNosav Wb
End Sub

Private Function ZTR_Fmtr() As LoFmtr
Dim O As LoFmtr
With O
    ReDim .Align(2)
    ReDim .AlignCno(2)
    
    .Align(0) = xlHAlignLeft
    .Align(1) = xlHAlignCenter
    .Align(2) = xlHAlignRight
    .AlignCno(0) = 1
    .AlignCno(1) = 2
    .AlignCno(2) = 3
End With
ZTR_Fmtr = O
End Function
