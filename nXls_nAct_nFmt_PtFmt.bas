Attribute VB_Name = "nXls_nAct_nFmt_PtFmt"
Option Compare Database
Option Explicit

Function PfsFny(A As PivotFields) As String()
PfsFny = ObjAyNy(A)
End Function

Sub PtFmt(A As PivotTable, F As PtFmtr)
Dim J%
ZSet_Pt A
ZSet_Ori_Row A, F.Row
ZSet_Ori_Col A, F.Col
ZSet_Ori_Pag A, F.Pag
ZSet_Ori_Dta A, F.Dta
ZSet_Fmt_Lbl A, F.LblColFld, F.LblDtaFno, F.LblVal
ZSet_Fmt_OutLin A, F.OutLinFno, F.OutLinLvl
ZSet_Tot_SubTot A, F.SubTotFno
ZSet_Tot_DtaSum A, F.DtaSumFno, F.DtaSumFun, F.DtaSumFmt
ZSet_Tot_GrandColTot A, F.GrandColTot, F.GrandColWdt
ZSet_Tot_GrandRowTot A, F.GrandRowTot
ZSet_Tot_OpnInd A, F.OpnInd
ZSet_Fmt_Wdt A, F.WdtFno, F.WdtVal
End Sub

Function PtWs(A As PivotTable) As Worksheet
Set PtWs = A.Parent
End Function

Function RgLasCol(Rg As Range) As Range
Set RgLasCol = RgC(Rg, RgNCol(Rg)).EntireColumn
End Function

Private Sub ZSet_Fmt_Lbl(A As PivotTable, ColFld$(), DtaFno%(), Lbl$())
Dim J%
Dim O As PivotField
For J = 0 To UB(ColFld)
    If ColFld(J) = "" Then
        Set O = A.DataFields(DtaFno(J))
    Else
        Set O = A.PivotFields(ColFld(J))
    End If
    O.Caption = Lbl(J)
Next
End Sub

Private Sub ZSet_Fmt_OutLin(A As PivotTable, Fno%(), Lvl() As Byte)
If A.LayoutRowDefault <> XlLayoutRowType.xlTabularRow Then Stop
Dim J%
PtWs(A).Outline.SummaryColumn = xlSummaryOnLeft
Dim O As PivotField
For J = 0 To UB(Fno)
    Set O = A.PivotFields(Fno(J))
    RgC(O.DataRange, 1).EntireColumn.OutlineLevel = Lvl(J)
Next
End Sub

Private Sub ZSet_Fmt_Wdt(A As PivotTable, Fno%(), Wdt%())
If A.LayoutRowDefault <> XlLayoutRowType.xlTabularRow Then Stop
Dim J%
Dim O As PivotField
For J = 0 To UB(Fno)
    Set O = A.PivotFields(Fno(J))
    RgC(O.DataRange, 1).ColumnWidth = Wdt(J)
Next
End Sub

Private Sub ZSet_Ori(A As PivotTable, Fld$(), Ori As XlPivotFieldOrientation)
Dim F As PivotField, J%
For J = 0 To UB(Fld)
    Set F = A.PivotFields(Fld(J))
    F.Orientation = Ori
    F.Position = J + 1
    If Ori = xlColumnField Or Ori = xlRowField Or Ori = xlPageField Then
        F.AutoSort xlAscending, Fld(J)
    End If
Next
End Sub

Private Sub ZSet_Ori_Col(A As PivotTable, Fld$())
ZSet_Ori A, Fld, XlPivotFieldOrientation.xlColumnField
End Sub

Private Sub ZSet_Ori_Dta(A As PivotTable, Fld$())
ZSet_Ori A, Fld, XlPivotFieldOrientation.xlDataField
End Sub

Private Sub ZSet_Ori_Pag(A As PivotTable, Fld$())
ZSet_Ori A, Fld, XlPivotFieldOrientation.xlPageField
End Sub

Private Sub ZSet_Ori_Row(A As PivotTable, Fld$())
ZSet_Ori A, Fld, XlPivotFieldOrientation.xlRowField
End Sub

Private Sub ZSet_Pt(A As PivotTable)
A.InGridDropZones = True
A.RowAxisLayout xlTabularRow
A.HasAutoFormat = False
End Sub

Private Sub ZSet_Tot_DtaSum(A As PivotTable, DtaFno%(), Sum() As XlConsolidationFunction, Fmt$())
If AyIsEmpty(DtaFno) Then Exit Sub
Dim J%
Dim O As PivotField
For J = 0 To UB(DtaFno)
    Set O = A.DataFields(DtaFno(J))
    O.Function = Sum(J)
    O.NumberFormat = Fmt(J)
Next
End Sub

Private Sub ZSet_Tot_GrandColTot(A As PivotTable, Tot As Boolean, Wdt%)
A.ColumnGrand = Tot
If Tot Then RgLasCol(A.DataBodyRange).ColumnWidth = Wdt
End Sub

Private Sub ZSet_Tot_GrandRowTot(A As PivotTable, Tot As Boolean)
A.RowGrand = Tot
End Sub

Private Sub ZSet_Tot_OpnInd(A As PivotTable, OpnInd As Boolean)
A.ShowDrillIndicators = Not OpnInd
A.ShowDrillIndicators = OpnInd
End Sub

Private Sub ZSet_Tot_SubTot(A As PivotTable, Fno%())
Dim J%
Dim O As PivotField
For J = 1 To A.PivotFields.Count
    Set O = A.PivotFields(J)
    If AyHas(Fno, J) Then
        O.Subtotals = Array(True, False, False, False, False, False, False, False, False, False, False, False)
    Else
        O.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End If
Next
End Sub
