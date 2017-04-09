Attribute VB_Name = "nXls_nFmt_RgFmtDef"
Option Compare Database
Option Explicit
Const C_Mod = "nXls_nFmt_RgFmtDef"
Public Type RgFmtDef
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

Sub AA5()
RgFmtDefSqBrk__Tst
End Sub

Function RgFmtDefFny() As String()
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
RgFmtDefFny = O
End Function

Function RgFmtDefHdrSq(A As RgFmtDef)
Dim DrAy()
DrAy = A.HdrDrAy
RgFmtDefHdrSq = DrAySq(DrAy)
End Function

Sub RgFmtDefHdrSq__Tst()
'1 Declare
Dim RptSq
Dim Act As RgFmtDef
Dim Exp As RgFmtDef

'2 Assign
Dim Fx$
Dim Wb As Workbook
    Fx = PthTstRes(C_Mod) & "RgFmtDefSqBrk__Tst.xlsx"
    Set Wb = FxWb(Fx)
RptSq = WsSq(Wb.Sheets(1))


'3 Calling
Act = RgFmtDefSqBrk(RptSq)
Dim OHdrSq
    OHdrSq = RgFmtDefHdrSq(Act)
SqBrw OHdrSq, "HdrSq"
'4 Asst
WbVis Wb
Stop
WbCls Wb, NoSav:=True
End Sub

Function RgFmtDefSqBrk(RptSq) As RgFmtDef
Dim O As RgFmtDef
With O
    Dim J%, A$, Dr()
    For J = 1 To UBound(RptSq, 1)
        Dr = SqDr(RptSq, J)
        A = Dr(0)
        If AyIsAllEmptyEle(AyRmvAt(Dr)) Then
            GoTo Nxt
        End If
        Select Case UCase(A)
        Case "FLD":          .Fld = Dr: RgFmtDefSqBrk = O: Exit Function '=====>>>
        Case "ISSEPLIN":     .IsSepLin = VarIsBoolTrue(Dr(1))
        Case "TOTONRIGHT":   .TotOnRight = VarIsBoolTrue(Dr(1))
        Case "TOTONBELOW":   .TotOnBelow = VarIsBoolTrue(Dr(1))
        Case "LVL":          .Lvl = Dr
        Case "WDT":          .Wdt = Dr
        Case "SUM":          .Sum = Dr
        Case "FORMULA":      .Formula = Dr
        Case "ALIGN":        .Align = Dr
        Case "FMT":          .NumFmt = Dr
        Case "VLIN":         .VLin = Dr
        Case "HDR"::          Push .HdrDrAy, Dr
        Case "LBL":          .Lbl = Dr
        End Select
Nxt:
    Next
End With
Er "Given {RptSq} does not have [Fld] in first column", 1
End Function

Sub RgFmtDefSqBrk__Tst()
'1 Declare
Dim RptSq
Dim Act As RgFmtDef
Dim Exp As RgFmtDef

'2 Assign
Dim Fx$
Dim Wb As Workbook
    Fx = PthTstRes(C_Mod) & "RgFmtDefSqBrk__Tst.xlsx"
    Set Wb = FxWb(Fx)
RptSq = WsSq(Wb.Sheets(1))


'3 Calling
Act = RgFmtDefSqBrk(RptSq)

'4 Asst
WbVis Wb
Stop
WbCls Wb, NoSav:=True
End Sub

Function RgFmtDefTpWs() As Worksheet
Dim O As Worksheet
Set O = WsNew
AyPutCellV RgFmtDefFny, WsA1(O)
Set RgFmtDefTpWs = O
End Function

Sub RgFmtDefTpWs__Tst()
Dim O As Worksheet
    Set O = RgFmtDefTpWs
    O.Application.Visible = True
Stop
End Sub
