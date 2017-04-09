Attribute VB_Name = "nXls_nFmt_ListObj"
Option Compare Database
Option Explicit

Sub AA()
ListObjFmt__Tst
End Sub

Sub ListObjFmt(A As ListObject, Fmtr As ListObjFmtr)
Dim O As ListObjFmtr
    O = Fmtr

Dim OHdrA1Cell As Range

Dim OWs As Worksheet
    Set OWs = A.Parent
    
Dim J&, C
'Align
For J = 0 To UB(O.Align)
    C = O.AlignCno(J)
    ListObjC(A, C).HorizontalAlignment = O.Align(J)
Next

'BackColr
For J = 0 To UB(O.BackColr)
    C = O.AlignCno(J)
    ListObjC(A, C).Interior.Color = O.BackColr(J)
Next

'Fld

'Font
For J = 0 To UB(O.FontColr)
    C = O.FontColrCno(J)
    ListObjC(A, C).Font.Color = O.FontColr(J)
Next

'Formula
For J = 0 To UB(O.Formula)
    C = O.FormulaCno(J)
    ListObjRC(A, 1, C).Formula = O.Formula(J)
Next

'HdrBackColr
For J = 0 To UB(O.HdrBackColr)
    C = O.HdrBackColrCno(J)
    RgRC(OHdrA1Cell, R, C).Interior.Color = O.HdrBackColr(J)
Next

'HdrFontColr
For J = 0 To UB(O.HdrFontColr)
    C = O.HdrFontColrCno(J)
    RgRC(OHdrA1Cell, R, C).Interior.Color = O.HdrFontColr(J)
Next


'IsSepLin
If O.IsSepLin Then RgSetSepLin A.DataBodyRange

'Lvl
For J = 0 To UB(O.LvlNo)
    C = O.LvlCno(J)
    ListObjEntireC(A, C).Outline = O.LvlNo(J)
Next

'NumFmt
For J = 0 To UB(O.NumFmt)
    C = O.NumFmtCno(J)
    ListObjC(A, C).NumberFormat = O.NumFmt(J)
Next
'=====================================================
Dim ColNm$

'SumAvg
For J = 0 To UB(O.SumAvgColNm)
    ColNm = O.SumAvgColNm(J)
    ListObjC(A, ColNm).TotalsCalculation = xlTotalsCalculationAverage
Next

'SumCnt
ColNm = O.SumCntColNm
ListObjC(A, ColNm).TotalsCalculation = xlTotalsCalculationCount

'SumTot
For J = 0 To UB(O.SumTotColNm)
    ColNm = O.SumTotColNm(J)
    ListObjC(A, ColNm).TotalsCalculation = xlTotalsCalculationSum
Next
'=====================================================
'SummaryRow
OWs.Outline.SummaryColumn = O.SummaryRow
OWs.Outline.SummaryColumn = O.SummaryCol
'=====================================================
If Not AyIsEmpty(O.VLinRightCno) Then
    For Each C In O.VLinRightCno
        RgSetVLinRight ListObjC(A, C)
    Next
End If
If Not AyIsEmpty(O.VLinLeftCno) Then
    For Each C In O.VLinLeftCno
        RgSetVLinLeft ListObjC(A, C)
    Next
End If

RgSetBdrAround A.DataBodyRange

Dim A1 As Range
    Set A1 = RgA1(A.DataBodyRange)
OWs.Activate
A1.Select
A1.Outline.ShowLevels 1, 1
End Sub

Function ListObjFmt__Tst()
Dim Wb As Workbook
Dim Ws As Worksheet
    Dim TstPth$
    Dim Fx$
    TstPth = PthTstRes("nXls_nFmt_Dt")
    Fx = TstPth & "DtPutCellWithFmt__Tst.xlsx"
    Set Wb = FxWb(Fx)
    Set Ws = WbFstWs(Wb)
    Wb.Application.Visible = True

Dim Cas%
For Cas = 1 To 3
    Dim Sq
    Dim R1, C1, R2, C2
'        Dim Ws As Worksheet
'        Set Ws = Wb.Sheets(Cas)
'        Sq = WsSq(Ws)
'        R1 = SqRIdx(Sq, "Fld", 1)
'        R2 = UBound(Sq, 1)
'        C1 = 2
'        C2 = UBound(Sq, 2)
    
    Dim Cell As Range
    Dim Dt As Dt
    Dim Fmtr As ListObjFmtr
    
        Dt = ResDt("")
        Set Cell = WsRC(Ws, 20, 1)
        Fmtr = WsFmtr(Ws)
    DtPutCellWithFmt Dt, Cell, Fmtr '<=========
Next
End Function
