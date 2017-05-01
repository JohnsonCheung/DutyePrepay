Attribute VB_Name = "nXls_nObj_nRg_nDo_Rg"
Option Compare Database
Option Explicit

Sub RgCpyFormula(Rg As Range)
'Aim: Copy formula at {Rg} download {pNRow} (including the row of {Rg}
Stop
'If pNRow <= 0 Then Exit Sub
'On Error GoTo R
'With Rg(1, 1)
'    .Formula = pFormula
'    .Copy
'End With
'Dim mWs As Worksheet: Set mWs = Rg.Parent
'mWs.Range(Rg(2, 1), Rg(pNRow, 1)).PasteSpecial xlPasteFormulas
'Exit Sub
'R: ss.R
'E: Set_Formula = True: ss.B cSub, cMod, "Rg,NRow,pFormula", RgToStr(Rg), pNRow, pFormula
End Sub

Function RgCpyFormulaByCmt(A As Range, pNRec&) As Boolean
'Aim: Assume the row above Rg contains formula of the row Rg in Cmt.
'     After setting the formula of each cell of row Rg, copy them downward until pRnoEnd
Const cSub$ = "RgCpyFormulaByCmt"
On Error GoTo R
Dim mCnoLas As Byte: If Fnd_CnoLas(mCnoLas, A) Then ss.A 1: GoTo E
On Error GoTo R
Dim iCno As Byte
Dim mWs As Worksheet: Set mWs = A.Parent
For iCno = A.Column To mCnoLas
    Dim mRgeFormula As Range: Set mRgeFormula = mWs.Cells(A.Row - 1, iCno)
    If IsCmt(mRgeFormula) Then
        Dim mFormula$: mFormula = mRgeFormula.Comment.Text
        If Set_Formula(mRgeFormula(2, 1), pNRec, mFormula) Then ss.A 3, "Error in set formula for column[" & iCno & "]": GoTo E
    End If
Next
A.Application.CutCopyMode = False
Exit Function
R: ss.R
E:
End Function

Sub RgCpyFormulaDn(A As Range)
'Aim:Copy formula from row Dta to row NmFld's comment
Dim Ws As Worksheet:
    Set Ws = Excel.Application.ActiveSheet
Dim R1&, R2&
    R1 = RgR1(A)
    R2 = RgR2(A)
Dim iCno%
Dim OToRg As Range
Dim OFormula$
For iCno = RgC1(A) To RgC2(A)
    Set OToRg = WsCRR(Ws, iCno, R1 + 1, R2)
    OFormula = OToRg(0, 1).Formula
    OToRg.Formula = OFormula
Next
End Sub

Sub RgMge(A As Range)
If A.Count = 1 Then Exit Sub
Dim V: V = A.Cells(1, 1).Value
A.Value = Null
A.Merge
A.Cells(1, 1).Value = V
End Sub

Sub RgMgeH(A As Range)

End Sub

Sub RgRplVal(Rg As Range, pFmVal$, pToVal$)
'Aim: Repl value in {Rg} from {pFmVal} to {pToVal}
Const cSub$ = "Repl_RgeVal"
On Error GoTo R
Rg.Application.DisplayAlerts = False

Dim mWs As Worksheet: Set mWs = Rg.Parent
mWs.Outline.ShowLevels 8, 8

Dim mCell As Range
Set mCell = Rg.Find(What:=pFmVal, LookIn:=xlValues _
    , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext _
    , MatchCase:=False, SearchFormat:=False)
While TypeName(mCell) <> "Nothing"
    mCell.Value = Replace(mCell.Value, pFmVal, pToVal)
    Set mCell = Rg.FindNext(mCell)
Wend
Rg.Application.DisplayAlerts = True
Exit Sub
R: ss.R
E:
End Sub

Sub RgRplVal__Tst()
Const cFfnFm$ = "R:\Sales Simulation\Simulation\Templates\Topaz Data Import file ({StreamCode}).xls"
Const cFfnTo$ = "c:\temp\a.xls"
Dim mWb As Workbook: If FxCpyAndOpn(mWb, cFfnFm, cFfnTo) Then Stop
Dim mWs As Worksheet: Set mWs = mWb.Sheets("SumTotalEuro {BrandGroupName}")
If Repl_RgeVal(mWs.Cells, "{BrandGroupName}", "Johnson") Then Stop
mWb.Application.Visible = True
mWs.Activate
End Sub

Sub RgSetBdrAround(A As Range, Optional LinSty As XlLineStyle = XlLineStyle.xlContinuous, Optional BdrWgt As XlBorderWeight = XlBorderWeight.xlMedium)
A.BorderAround LinSty, BdrWgt
Dim Rg As Range
Set Rg = RgC(A, A.Columns.Count + 1)
With Rg.Borders(xlEdgeTop)
    .LineStyle = LinSty
    .Weight = BdrWgt
End With

Set Rg = RgR(A, A.Rows.Count + 1)
With Rg.Borders(xlEdgeLeft)
    .LineStyle = LinSty
    .Weight = BdrWgt
End With
End Sub

Sub RgSetColOutLinColr(A As Range, pNLvl As Byte)
'Dim mRge As Range
'Dim mWs As Worksheet: Set mWs = Rg.Worksheet
''-- Find mAyRgeRno(): Use first column & pNLvl & color in the cells to find which RgeRno of cells needs to set color
'Dim mAyRgeRno() As tRgeRno, mN%
'mN = 0
'Dim iRno&
'Dim mColrWhite&:  mColrWhite = 16777215
'Dim mRnoFm&: mRnoFm = 0
'For iRno = Rg.Row + pNLvl To Rg.Row + Rg.Rows.Count - 1
'    Set mRge = mWs.Cells(iRno, Rg.Column)
'    If mRge.Interior.Color = mColrWhite Then
'        If mRnoFm <> 0 Then
'            ReDim Preserve mAyRgeRno(mN)
'            mAyRgeRno(mN).Fm = mRnoFm
'            mAyRgeRno(mN).To = iRno - 1
'            mN = mN + 1
'            mRnoFm = 0
'        End If
'        GoTo Nxt
'    End If
'    If mRnoFm = 0 Then mRnoFm = iRno
'Nxt:
'Next
'If mRnoFm <> 0 Then
'    ReDim Preserve mAyRgeRno(mN)
'    mAyRgeRno(mN).Fm = mRnoFm
'    mAyRgeRno(mN).To = iRno - 1
'End If
''-- Find mAyColr(1-pNLvl) by Rg downward pNLvl cells
'ReDim mAyColr&(pNLvl)
'Dim iLvl As Byte
'For iLvl = 1 To pNLvl - 1
'    Set mRge = Rg.Cells(iLvl, 1)
'    mAyColr(iLvl) = mRge.Interior.Color
'Next
''-- Loop all col and set color
'Dim iCno As Byte
'For iCno = Rg.Column To Rg.Column + Rg.Columns.Count - 1
'    iLvl = mWs.Columns(iCno).OutlineLevel
'    If iLvl < pNLvl Then
'        Dim mColr&: mColr = mAyColr(iLvl)
'        '--
'        Set mRge = mWs.Range(mWs.Cells(Rg.Row + iLvl, iCno), mWs.Cells(Rg.Row + pNLvl - 1, iCno))
'        mRge.MergeCells = True
'        mRge.Interior.Color = mColr
'        mRge.Borders(xlEdgeTop).LineStyle = xlLineStyleNone
'        '
'        Dim J As Byte
'        For J = 0 To UBound(mAyRgeRno)
'            Set mRge = mWs.Range(mWs.Cells(mAyRgeRno(J).Fm, iCno), mWs.Cells(mAyRgeRno(J).To, iCno))
'            mRge.Interior.Color = mColr
'        Next
'    End If
'Next
End Sub

Sub RgSetColOutLine(A As Range, nCol%, NLvl%)
'Aim: use pRno:pCnoBeg-pCnoEnd to set column outline.
'If IsNothing(A) Then Er "RgSetColOutLin: Given-{Rg} is nothing"
'Dim iCno%, iRno&
'Dim mWs As Worksheet: Set mWs = Rg.Parent
'Dim mRge As Range
'Dim iLvl As Byte
'For iCno = Rg.Column + 1 To Rg.Column + pNCol - 1
'    For iLvl = 2 To pNLvl
'        iRno = Rg.Row + iLvl
'        If IsEmpty(mWs.Cells(iRno, iCno).Value) Then
'            Set mRge = mWs.Cells(1, iCno)
'            Set mRge = mRge.EntireColumn
'            mRge.OutlineLevel = iLvl
'            GoTo NxtCol
'        End If
'    Next
'    Set mRge = mWs.Cells(1, iCno)
'    Set mRge = mRge.EntireColumn
'    mRge.OutlineLevel = iLvl
'NxtCol:
'Next
'Dim mAyRgeCno() As tRgeCno
'For iLvl = 0 To 2
'    iRno = iLvl + Rg.Row
'    Set mRge = mWs.Range(mWs.Cells(iRno, Rg.Column), mWs.Cells(iRno, Rg.Column + pNCol - 1))
'    If Fnd_AyRgeCno(mAyRgeCno, mRge) Then GoTo E
'    Dim J%
'    For J = 0 To UBound(mAyRgeCno)
'        With mAyRgeCno(J)
'            MgeRge mWs.Range(mWs.Cells(iRno, .Fm), mWs.Cells(iRno, .To))
'        End With
'    Next
'Next
End Sub

Sub RgSetHypLnkToWs(A As Excel.Range)
'Aim: Set any cells within the {Rg} to hyper link to A1 of worksheet if they have the same value
Dim Ws As Worksheet: Set Ws = A.Worksheet
Dim Wb As Workbook: Set Wb = Ws.Parent
Dim WsNy$(): WsNy = WbWsNy(Wb)
Dim Cell As Range, V, J%
For Each Cell In A
    V = Cell.Value
    If VarType(V) = vbString Then
        V = Left(V, 31)
        If AyHas(WsNy, V) Then A.Hyperlinks.Add Cell, "", FmtQQ("'?'!A1", V)
    End If
Next
End Sub

Sub RgSetHypLnkToWs__Tst()
Dim Ws As Worksheet: Set Ws = WsNew("Index")
Dim Wb As Workbook: Set Wb = Ws.Parent
Dim I As Worksheet, J%
For J = 1 To 100
    Set I = Wb.Sheets.Add
    I.Name = "Ws-" & J
Next
For J = 2 To 100
    Ws.Range("A" & J).Value = "Ws-" & J
Next

RgSetHypLnkToWs Ws.Range("A1:A100")
WbVis Wb
Stop
WbClsNosav Wb
End Sub

Sub RgSetNmH(A As Range)
Dim OKeyAdrDic As Dictionary
    Set OKeyAdrDic = RgKeyAdrDicH(A)
    If DicIsEmpty(OKeyAdrDic) Then Exit Sub
Dim oWs As Worksheet
    Set oWs = RgWs(A)
    Dim J%
    Dim Nm$, Adr$, K
    For Each K In OKeyAdrDic
        Nm = K
        Adr = OKeyAdrDic(K)
        'OWs.Names.Add "x" & mWs.Cells(G.cRnoDta - 1, ICno).Value, mWs.Columns(ICno)
        oWs.Names.Add Nm, oWs.Range(Adr)    '<====
    Next
End Sub

Sub RgSetNoWrap(A As Range)
A.WrapText = False
End Sub

Sub RgSetSepLin(A As Range, Optional NLin% = 3)
'[SepLin] Set Separating Lines
Dim ONoNeed As Boolean
    Stop
    
If ONoNeed Then Exit Sub
Dim ORg1 As Range, ORg2 As Range
Dim oRno&
    Dim R2&
        Stop
    Set ORg1 = RgRR(A, 1, NLin)
    Set ORg2 = RgRR(A, NLin + 1, R2)

RgR(A, oRno).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlDot
    
ORg1.Copy
ORg2.PasteSpecial xlPasteFormats
End Sub

Sub RgSetSumHLeft(HRg As Range)
Dim A1$, A2$
    Stop
Const C = "=Sum(?:?)"
Dim F$: F = FmtQQ(C, A1, A2)
RgA1(HRg).Formula = F
End Sub

Sub RgSetVdt(A As Range, pLv$ _
    , Optional InpTit$ = "Enter value or leave blank" _
    , Optional InpMsg$ = "Enter one of the value in the list or leave it blank." _
    , Optional ErTit$ = "Not in the List" _
    , Optional ErMsg = "Please enter a value in list or leave it blank" _
    )
'Aim: Set the validation of {Rg} to select a list of value {pLv}.
'     'The list of value' will be the stored in the avaliable column of ws [SelectionList]
'          Ws [SelectionList] Row1=Ws Name, Row2=Rge that will use to the list to select value, Row3 and onward will be the selection value
' Do Build mRgeLv: 'The list of value'
Dim mRgeLv As Range:  Set_Lv2ColAtEnd mRgeLv, pLv, A.Worksheet, A.Address
' Do Set Vdt of Rg
Do
    With A.Validation
        .Delete
        Dim mFormula$: mFormula = "=" & mRgeLv.Address
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=mFormula
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = InpTit
        On Error Resume Next
        .InputMessage = InpMsg
        .ErrorTitle = ErTit
        .ErrorMessage = ErMsg
        .ShowInput = True
        .ShowError = True
    End With
Loop Until True
End Sub

Function RgSetVdt__Tst()
Const cSub$ = "RgSetVdt_Tst"
Dim mRge As Range, mLv$
Dim mRslt As Boolean, mCase As Byte: mCase = 1

Dim mWb As Workbook: If Crt_Wb(mWb, "c:\aa.xls", True) Then ss.A 1: GoTo E
mWb.Application.Visible = True
Select Case mCase
Case 1
    Set mRge = mWb.Sheets(1).Range("A1:D5")
    mLv = "aa,bb,cc,11,22,33"
End Select
RgSetVdt mRge, mLv
Exit Function
R: ss.R
E:
End Function

Sub RgSetVLinLeft(A As Range)
BdrSet_Continuous_Medium A.Borders(xlEdgeLeft)
End Sub

Sub RgSetVLinRight(A As Range)
BdrSet_Continuous_Medium A.Borders(xlEdgeRight)
End Sub

Function RsPutCell(A As DAO.Recordset, Cell As Range) As Range
Set RsPutCell = SqPutCell(RsSq(A), Cell)
End Function

Function RsPutCell__Tst()
TblCrt_ByFldDclStr "#Tmp", "aa text 10, bb memo"
With CurrentDb.TableDefs("#Tmp").OpenRecordset
    Dim J%
    For J = 0 To 10
        .AddNew
        !aa = J
        !BB = String(245, Chr(Asc("0") + J))
        .Update
    Next
    For J = 0 To 10
        .AddNew
        !aa = J
        !BB = String(500, Chr(Asc("0") + J))
        .Update
    Next
    .Close
End With
Dim Ws As Worksheet: Set Ws = WsNew
Dim Cell As Range: Set Cell = WsA1(Ws)
Dim Rs As Recordset: Set Rs = SqlRs("Select * from [#Tmp]")
RsPutCell Rs, Cell
Cell.Application.Visible = True
End Function

