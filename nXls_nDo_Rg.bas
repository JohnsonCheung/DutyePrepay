Attribute VB_Name = "nXls_nDo_Rg"
Option Compare Database
Option Explicit

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
E: RgCpyFormulaByCmt = True: ss.B cSub, cMod, "Rg,pNRec", ToStr_Rge(A)
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

Sub RgSetColOutLine(A As Range, NCol%, NLvl%)
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

Sub RgSetNmH(A As Range)
Dim OKeyAdrDic As Dictionary
    Set OKeyAdrDic = RgKeyAdrDicH(A)
    If DicIsEmpty(OKeyAdrDic) Then Exit Sub
Dim OWs As Worksheet
    Set OWs = RgWs(A)
    Dim J%
    Dim Nm$, Adr$, K
    For Each K In OKeyAdrDic
        Nm = K
        Adr = OKeyAdrDic(K)
        'OWs.Names.Add "x" & mWs.Cells(G.cRnoDta - 1, ICno).Value, mWs.Columns(ICno)
        OWs.Names.Add Nm, OWs.Range(Adr)    '<====
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
Dim ORno&
    Dim R2&
        Stop
    Set ORg1 = RgRR(A, 1, NLin)
    Set ORg2 = RgRR(A, NLin + 1, R2)

RgR(A, ORno).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlDot
    
ORg1.Copy
ORg2.PasteSpecial xlPasteFormats
End Sub

Sub RgSetVLinLeft(A As Range)
BdrSetVLin A.Borders(xlEdgeLeft)
End Sub

Sub RgSetVLinRight(A As Range)
BdrSetVLin A.Borders(xlEdgeRight)
End Sub

Function RsCpyToFrm(pRs As DAO.Recordset, pFrm As Access.Form, FnStr$) As Boolean
'Aim: Copy the fields value from {pRs} to the controls in {pFrm}.  Only those fields in {FnStr} will be copied.
'     {FnStr} is in fmt of aaa=xxx,bbb,ccc  aaa,bbb,ccc will be field name in {pFrm} & xxx,bbb,ccc will be field in {pRs}
Const cSub$ = "RsCpyToFrm"
On Error GoTo R
Dim mAn_Frm$(), mAn_Rs$(): If Brk_Lm_To2Ay(mAn_Frm, mAn_Rs, FnStr) Then ss.A 1: GoTo E
Dim mIsEq As Boolean, mEr$, mV_Rs, mV_FrmNew
Dim J%: For J = 0 To Siz_Ay(mAn_Frm) - 1
    With pFrm.Controls(mAn_Frm(J))
        mV_Rs = pRs.Fields(mAn_Rs(J)).Value
        mV_FrmNew = .Value
        If IfEq(mIsEq, mV_Rs, mV_FrmNew) Then ss.A 1: GoTo E
        If Not mIsEq Then .Value = mV_Rs
    End With
Next
'Sav.Rec
Exit Function
R: ss.R
E: RsCpyToFrm = True: ss.B cSub, cMod, "pRs,pFrm,FnStr", ToStr_Rs(pRs), ToStr_Frm(pFrm), FnStr
End Function

Function RsPutCell(Rs As Recordset, Cell As Range) As Boolean
DtPutCell RsDt(Rs), Cell
End Function

Function RsPutCell__Tst()
TblCrt_ByFldDclStr "#Tmp", "aa text 10, bb memo"
With CurrentDb.TableDefs("#Tmp").OpenRecordset
    Dim J%
    For J = 0 To 10
        .AddNew
        !AA = J
        !BB = String(245, Chr(Asc("0") + J))
        .Update
    Next
    For J = 0 To 10
        .AddNew
        !AA = J
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
