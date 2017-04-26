Attribute VB_Name = "nXls_nFmt_LoFmtr"
Option Compare Database
Option Explicit
Const C_Mod$ = "nXls_nFmt_LOFmtr"
Private Type ZAlign
    Align() As XlHAlign
    AlignCno() As Integer
End Type
Private Type ZColr
    Colr() As Long
    Cno() As Integer
End Type
Private Type ZFormula
    Formula() As String
    Cno() As Integer
End Type
Private Type ZHdrColr
    Rno() As Integer
    Cno() As Integer
    Colr() As Long
End Type
Private Type ZVLin
    Cno() As Integer
End Type
Type LoFmtr
    HdrSq() As Variant
    
    VLinLeftCno() As Integer
    VLinRightCno() As Integer
    
    IsSepLin As Boolean
    SummaryCol As XlSummaryColumn
    SummaryRow As XlSummaryRow
    
    FormulaCno() As Integer
    Formula() As String
    
    FontColrCno() As Integer
    FontColr() As Long
    
    BackColrCno() As Integer
    BackColr() As Long
    
    SumTotColNm() As String
    SumAvgColNm() As String
    SumCntColNm As String
    
    LvlCno() As Integer
    LvlNo() As Integer
    
    HdrFontColrCno() As Integer
    HdrFontColrRno() As Integer
    HdrFontColr() As Long
    HdrBackColrCno() As Integer
    HdrBackColrRno() As Integer
    HdrBackColr() As Long
    
    AlignCno() As Integer
    Align() As XlHAlign
    
    NumFmtCno() As Integer
    NumFmt() As String
    
    Lbl() As String
    Fld() As String
End Type
Private Type ZFmtrLin
    IsSepLin As Boolean     'IsSepLin;
    TotOnRight As Boolean    'TotOnRight;
    TotOnBelow As Boolean     'TotOnBelow;
    Lvl() As Variant         'Lvl;L1;L2;L3;L4
    Sum() As Variant         'Sum;Tot;Sum;Cnt
    Formula() As Variant  'Formula;XX;XX
    Align() As Variant  'Align;L;R;;L
    NumFmt() As Variant     'Fmt;XX;XX;;XX
    VLin() As Variant    'VLin;L;R;LR;
    Lbl() As Variant      'Lbl;X;X;X
    HdrDrAy() As Variant    'H;XX;XX
    Fld() As Variant       'Fld;X;X
    BackColr() As Variant
    FontColr() As Variant
    HdrBackColrDrAy() As Variant
    HdrFontColrDrAy() As Variant
    Wdt() As Variant
End Type

Sub AA00()
LOFmtr__Tst
End Sub

Sub AA5()
ZFmtrSqBrk__Tst
End Sub

Function ListObjCrt(A As Range) As ListObject
Set ListObjCrt = RgWs(A).ListObjects.Add(xlSrcRange, A)
End Function

Sub ListObjFmt(A As ListObject, Fmtr As LoFmtr)
Dim Rg As Range
Dim J&, C, R%
'=========================
Dim O As LoFmtr
    O = Fmtr

Dim OHdrA1Cell As Range

Dim OWs As Worksheet
    Set OWs = A.Parent
    
'Align ====================
For J = 0 To UB(O.Align)
    C = O.AlignCno(J)
    ListObjC(A, C).HorizontalAlignment = O.Align(J)
Next

'BackColr =================
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
    R = O.HdrFontColrRno(J)
    RgRC(OHdrA1Cell, R, C).Interior.Color = O.HdrFontColr(J)
Next


'IsSepLin
If O.IsSepLin Then RgSetSepLin A.DataBodyRange

'Lbl

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
If ColNm <> "" Then
    ListObjC(A, ColNm).TotalsCalculation = xlTotalsCalculationCount
End If

'SumTot
For J = 0 To UB(O.SumTotColNm)
    ColNm = O.SumTotColNm(J)
    ListObjC(A, ColNm).TotalsCalculation = xlTotalsCalculationSum
Next
'=========================
'SummaryRow/Col
Select Case O.SummaryCol
Case xlSummaryOnLeft, xlSummaryOnRight: OWs.Outline.SummaryColumn = O.SummaryCol
End Select
Select Case O.SummaryRow
Case xlSummaryAbove, xlSummaryBelow: OWs.Outline.SummaryRow = O.SummaryRow
End Select
'=========================
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
'=========================
    Set Rg = A.DataBodyRange
If Not IsNothing(Rg) Then
    RgSetBdrAround Rg
    CellAct Rg
End If

OWs.Activate
OWs.Outline.ShowLevels 1, 1
End Sub

Sub LoFmt(A As ListObject, F As LoFmtr)

End Sub

Function LoFmtr(FmtrSq) As LoFmtr
Dim Def As ZFmtrLin
    Def = ZFmtrSqBrk(FmtrSq)
Dim J%
Dim O As LoFmtr
'Align
With ZAlign(Def.Align)
    O.Align = .Align
    O.AlignCno = .AlignCno
End With

'BackColr
With ZColr(Def.BackColr)
    O.BackColr = .Colr
    O.BackColrCno = .Cno
End With

'Fld
O.Fld = AyRmvAt(Def.Fld)
    
'FontColr
With ZColr(Def.FontColr)
    O.FontColrCno = .Cno
    O.FontColr = .Colr
End With

'Formula
With ZFormula(Def.Formula)
    O.FormulaCno = .Cno
    O.Formula = .Formula
End With

'HdrBackColr
With ZHdrColr(Def.HdrFontColrDrAy)
    O.HdrBackColr = .Colr
    O.HdrBackColrCno = .Cno
    O.HdrBackColrRno = .Rno
End With

'HdrFontColr
With ZHdrColr(Def.HdrFontColrDrAy)
    O.HdrFontColr = .Colr
    O.HdrFontColrCno = .Cno
    O.HdrFontColrRno = .Rno
End With
X:
LoFmtr = O
End Function

Function LOFmtr__Tst()
Dim Wb As Workbook
    Set Wb = ZTR_Wb
    WbVis Wb
Dim Cas%
For Cas = 1 To 3
    Dim DtSq
    Dim Ws As Worksheet
        Dim R1, C1, R2, C2, Sq
        Set Ws = Wb.Sheets(Cas)
        Sq = WsSq(Ws)
        R1 = SqRIdx(Sq, "Fld", 1)
        R2 = UBound(Sq, 1)
        C1 = 2
        C2 = UBound(Sq, 2)
        DtSq = WsRCRC(Ws, R1, C1, R2, C2).Value
    Dim Cell As Range
    Dim Dt As Dt
    Dim Fmtr As LoFmtr
    
        Dt = DtNewSq(DtSq)
        Set Cell = WsRC(Ws, 20, 1)
        Fmtr = WsFmtr(Ws)
    DtPutCellWithFmt Dt, Cell, Fmtr '<=========
Next
Stop
WbClsNosav Wb
End Function

Function LOFmtrToLy(A As LoFmtr) As String()

End Function

Sub LOFmtrTp__Tst()
Dim O As Worksheet
    Set O = LOFmtrTp
    O.Application.Visible = True
Stop
End Sub

Sub ZAlign__Tst()
Dim A()
Dim Act As ZAlign
    
A = Array("Align", "LR", "L", "R", "C", 1, Empty)
Act = ZAlign(A)

AyAsstEqExa Act.Align, Array(xlHAlignLeft, xlHAlignRight, xlHAlignCenter)
AyAsstEqExa Act.AlignCno, Array(2, 3, 4)
End Sub

Sub ZColr__Tst()
Dim A()
Dim Act As ZColr
    
A = Array("BackColr", 1232#, , 12321#, 12222#, Empty)
Act = ZColr(A)

AyAsstEqExa Act.Colr, Array(1232&, 12321&, 12222&)
AyAsstEqExa Act.Cno, Array(1, 3, 4)

End Sub

Sub ZDefHdrSq__Tst()
'1 Declare
Dim FmtrSq
Dim Act As ZFmtrLin
Dim Exp As ZFmtrLin

'2 Assign
FmtrSq = ZTR_FmtrSq

'3 Calling
Act = ZFmtrSqBrk(FmtrSq)
Dim OHdrSq
    OHdrSq = ZDefHdrSq(Act)
SqBrw OHdrSq, "HdrSq"
'4 Asst
Stop
End Sub

Sub ZFmtrSqBrk__Tst()
'1 Declare
Dim FmtrSq
Dim Act As ZFmtrLin
Dim Exp As ZFmtrLin

'2 Assign
FmtrSq = ZTR_FmtrSq

'3 Calling
Act = ZFmtrSqBrk(FmtrSq)

'4 Asst
Stop
End Sub

Sub ZHdrColr__Tst()
Dim A()
Dim Act As ZHdrColr
    
Push A, Array("BackColr", 1232#, , 12321#, 12222#, Empty)
Push A, Array("BackColr", 1232#, , 12321#, 12222#, Empty)
Push A, Array("BackColr", 1232#, , 12321#, 12222#, Empty)
Act = ZHdrColr(A)

AyAsstEqExa Act.Colr, Array(1232&, 12321&, 12222&, 1232&, 12321&, 12222&, 1232&, 12321&, 12222&)
AyAsstEqExa Act.Rno, Array(0, 0, 0, 1, 1, 1, 2, 2, 2)
AyAsstEqExa Act.Cno, Array(1, 3, 4, 1, 3, 4, 1, 3, 4)
End Sub

Private Function LOFmtrTp() As Worksheet
Dim O As Worksheet
Set O = WsNew
AyPutCellV ZDefFny, WsA1(O)
Set LOFmtrTp = O
End Function

Private Function ZAlign(AlignDr()) As ZAlign
Dim O As ZAlign
Dim Align$, J%
For J = 1 To UB(AlignDr)
    Align = UCase(AlignDr(J))
    If Len(Align) = 1 Then
        Select Case Align
        Case "L"
            Push O.AlignCno, J
            Push O.Align, XlHAlign.xlHAlignLeft
        Case "R"
            Push O.AlignCno, J
            Push O.Align, XlHAlign.xlHAlignRight
        Case "C"
            Push O.AlignCno, J
            Push O.Align, XlHAlign.xlHAlignCenter
        End Select
    End If
Next
ZAlign = O
End Function

Private Function ZColr(ColrDr()) As ZColr
Dim O As ZColr
Dim J%
For J = 1 To UB(ColrDr)
    If VarIsDbl(ColrDr(J)) Then
        Push O.Cno, J
        Push O.Colr, ColrDr(J)
    End If
Next
ZColr = O
End Function

Private Function ZDefFny() As String()
Dim O$()
Push O, "IsSepLin"
Push O, "TotOnRight"
Push O, "TotOnBelow"
Push O, "Lvl"
Push O, "Wdt"
Push O, "Sum"
Push O, "VLin"
Push O, "Hdr"
Push O, "Lbl"
Push O, "ColrFont"
Push O, "ColrBack"
Push O, "Fld"
ZDefFny = O
End Function

Private Function ZDefHdrSq(A As ZFmtrLin)
Dim DrAy()
DrAy = A.HdrDrAy
ZDefHdrSq = DrAySq(DrAy)
End Function

Private Function ZFmtrSqBrk(FmtrSq) As ZFmtrLin
Dim O As ZFmtrLin
With O
    Dim J%, A$, Dr()
    For J = 1 To UBound(FmtrSq, 1)
        Dr = SqDr(FmtrSq, J)
        A = Dr(0)
        If AyIsAllEmptyEle(AyRmvAt(Dr)) Then
            GoTo Nxt
        End If
        Select Case A
        Case "Fld":          .Fld = Dr: ZFmtrSqBrk = O: Exit Function '=====>>>
        Case "IsSepLin":     .IsSepLin = VarIsBoolTrue(Dr(1))
        Case "TotOnRight":   .TotOnRight = VarIsBoolTrue(Dr(1))
        Case "TotOnBelow":   .TotOnBelow = VarIsBoolTrue(Dr(1))
        Case "Lvl":          .Lvl = Dr
        Case "Wdt":          .Wdt = Dr
        Case "Sum":          .Sum = Dr
        Case "Formula":      .Formula = Dr
        Case "Align":        .Align = Dr
        Case "NumFmt":       .NumFmt = Dr
        Case "VLin":         .VLin = Dr
        Case "Hdr"::          Push .HdrDrAy, Dr
        Case "Lbl":          .Lbl = Dr
        End Select
Nxt:
    Next
End With
Er "Given {FmtrSq} does not have [Fld] in first column", 1
End Function

Private Function ZFormula(FormulaDr()) As ZFormula
Dim J%, Formula$, O As ZFormula
For J = 1 To UB(FormulaDr)
    Formula = FormulaDr(J)
    Push O.Cno, J
    Push O.Formula, Formula
Next
ZFormula = O
End Function

Private Sub ZFormula__Tst()
Dim A()
Dim Act As ZFormula
    
A = Array("BackFormula", "sddsfsdf", , "lskdfsdlkf", "slkdfjsdf", Empty)
Act = ZFormula(A)

AyAsstEqExa Act.Formula, ApSy("sddsfsdf", "lskdfsdlkf", "slkdfjsdf")
AyAsstEqExa Act.Cno, Array(1, 3, 4)
End Sub

Private Function ZHdrColr(HdrColrDrAy()) As ZHdrColr
Dim O As ZHdrColr
Dim R&, Dr, J%
If Not AyIsEmpty(HdrColrDrAy) Then
    R = -1
    For Each Dr In HdrColrDrAy
        R = R + 1
        For J = 1 To UB(Dr)
            If VarIsDbl(Dr(J)) Then
                Push O.Cno, J
                Push O.Rno, R
                Push O.Colr, Dr(J)
                
            End If
        Next
        R = R + 1
    Next
End If
ZHdrColr = O
End Function

Private Sub ZTA_ResWbBrw()
WbVis ZTR_Wb
End Sub

Private Sub ZTA_TstResPthBrw()
PthTstResBrw C_Mod
End Sub

Private Sub ZTA_WbBrw()
WbVis ZTR_Wb
End Sub

Private Function ZTR_Dt() As Dt
ZTR_Dt = DtNewSclVBar("Tbl;ABC|Fld;A;B;C;D|;1;2;3;4|;A;B;C;D")
End Function

Private Function ZTR_FmtrSq()
Static X As Boolean, Sq
If Not X Then
    X = True
    Dim Wb As Workbook
    Set Wb = ZTR_Wb
    Sq = WsSq(Wb.Sheets(1))
    WbCls Wb, NoSav:=True
End If
ZTR_FmtrSq = Sq
End Function

Private Sub ZTR_FmtrSq__Tst()
SqBrw ZTR_FmtrSq
End Sub

Private Function ZTR_Wb() As Workbook
Dim Fx$
    Fx = PthTstRes(C_Mod) & "Res.xlsx"
Set ZTR_Wb = FxWb(Fx)
End Function

Private Function ZTR_Wb1() As Workbook
Dim TstPth$
Dim Fx$
    TstPth = PthTstRes("nXls_nFmt_Dt")
    Fx = TstPth & "DtPutCellWithFmt__Tst.xlsx"

Set ZTR_Wb = FxWb(Fx)
End Function
