Attribute VB_Name = "nXls_nObj_nWs_nDo_Ws"
Option Compare Database
Option Explicit

Function WbAddOrRplWs(A As Workbook, WsNm$) As Worksheet
If WbHasWs(A, WsNm) Then
    A.Sheets.Add
Else
End If

End Function

Sub WsAddContent(A As Worksheet, FmSngWsFx$)
Dim Dt As Dt
Dt = SngWsFxDt(FmSngWsFx)
WsAddContentFmDt A, Dt
End Sub

Function WsAddContent__Tst()
Const cFfn1$ = "c:\tmp\a.xls"
Dim Ws As Worksheet: Set Ws = WsNew
WsAddContent Ws, cFfn1
Ws.Application.Visible = True
End Function

Sub WsAddContentFmDt(A As Worksheet, Dt As Dt)
'Aim: Add the content of {FxFm} at the end of {pWsTo} provided that they are same layout
'Assume: [1 Ws in FxFm] & [Same Layout]
''[1 Ws in FxFm] There is only one ws in {FxFm} having the same name of the file name of {FxFm}
''[Same Layout]       The column headings of {FxFm} & {pWsTo} should be the same
'==Start
Dim J As Byte
For J = 1 To 254
    Dim mV: mV = A.Cells(1, J).Value
    If mV = "" Then Exit Sub
    If mV <> A.Cells(1, J).Value Then Er "Err"
Next

'Copy
Dim AdrFm$, AdrTo$
Dim mColToLast$: mColToLast = A.Cells.SpecialCells(xlCellTypeLastCell).Column
AdrFm = "A2:" & A.Cells.SpecialCells(xlCellTypeLastCell).Address
AdrTo = "A" & A.Cells.SpecialCells(xlCellTypeLastCell).Row + 1
A.Range(AdrFm).Copy
A.Range(AdrTo).PasteSpecial xlPasteValues
End Sub

Sub WsClrOutLine(A As Worksheet)
'Aim: Clear all columns' outline
Dim J As Byte
For J = 1 To 9
    On Error GoTo X
    A.Cells.EntireColumn.Ungroup
Next
X:
End Sub

Sub WsCpy(A As Worksheet, ToWb As Workbook, Optional ToWsNm$)
'Aim: Copy {pWbFm}!{pWsNmFm$} to {pWbTo}.  If ws exist in {pWbTo}, ws will be replaced and position retended, else copy to end
Dim OToAftWs As Worksheet
    Dim ToWNm$: ToWNm = StrNz(ToWsNm, A.Name)

A.Copy , OToAftWs
End Sub

Function WsCpy__Tst()
Dim mFxFm$: mFxFm = "c:\tmp\Fm.xls"
Dim mFxTo$: mFxTo = "c:\tmp\To.xls"
Dim mWbFm As Workbook: If Crt_Wb(mWbFm, mFxFm, True) Then Stop
Dim mWbTo As Workbook: If Crt_Wb(mWbTo, mFxTo, True) Then Stop
Dim mWsFm As Worksheet: Set mWsFm = mWbFm.Sheets.Add
If Set_Ws_ByLpAp(mWsFm, 1, 1, False, "Msg", "First Time From Wb") Then Stop
mWbFm.Application.Visible = True
MsgBox "Before Copy, Check To ws"
Stop
WsCpy mWbTo, mWbFm, mWsFm.Name
Stop
If Set_Ws_ByLpAp(mWsFm, 1, 1, False, "Msg", "Second Time From Wb") Then Stop
WsCpy mWbTo, mWbFm, mWsFm.Name
MsgBox "Second Time Copy.  Check To ws"
End Function

Function WsCpy1(oWsTar As Worksheet, pWsFm As Worksheet, Optional pWsNmTo$ = "", Optional pWbTo As Workbook = Nothing) As Boolean
Const cSub$ = "WsCpy"
'Aim: Copy {pWsFm} to a new {oWsTar}.
'Note: If {pWbTo} is given, the new Ws will be at end of {pWbTo}.  Otherwise, the oWsTar will be the same workbook as pWsFm.
'Note: If {pWsNmTo} is not given, the new Ws Name will use the {pWsFm}
If pWsNmTo = "" And IsNothing(pWbTo) Then ss.A 1, "CpyWs must be given Either or both of {pWsNmTo}, {pWbTo}": GoTo E
'==Start
'Set {mWbTo} & {mWsNmTo}
Dim mWbTo As Workbook, mWsNmTo$
If IsNothing(pWbTo) Then
    Set mWbTo = pWsFm.Parent
Else
    Set mWbTo = pWbTo
End If
If pWsNmTo = "" Then
    mWsNmTo = pWsFm.Name
Else
    mWsNmTo = pWsNmTo
End If
'Copy and Paste
pWsFm.Cells.Copy
Dim mWsNew As Worksheet: If Add_Ws(mWsNew, mWbTo, mWsNmTo) Then ss.A 2: GoTo E
mWsNew.Cells.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
mWsNew.Cells.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Set oWsTar = mWsNew
Exit Function
R: ss.R
E:
End Function

Function WsCpyRow(pWs As Worksheet, pRnoFm&, pRnoTo&) As Boolean
'Aim: Copy {pRnoFm} to {pRnoTo} in pWs
Const cSub$ = "WsCpyRow"
On Error GoTo R
Dim mRowFm As Range: Set mRowFm = pWs.Range(pRnoFm & ":" & pRnoFm)
Dim mRowTo As Range: Set mRowTo = pWs.Range(pRnoTo & ":" & pRnoTo)
mRowFm.Copy
mRowTo.PasteSpecial xlPasteAllExceptBorders
pWs.Application.CutCopyMode = False
Exit Function
R: ss.R
E:
End Function

Function WsCpyRowDown(pWs As Worksheet, pColRgeList$, pRow&, pNRow&, Optional pCopyOnly_Val_Fmt As Boolean = True) As Boolean
Dim mAyColRge$(), J As Byte, mFmAdr$, mToAdr$
mAyColRge = Split(pColRgeList, CtComma)
With pWs
    For J = LBound(mAyColRge) To UBound(mAyColRge)
        mFmAdr = Cv_RnoColRge2Adr(mAyColRge(J), pRow)
        mToAdr = Cv_RnoColRge2Adr(mAyColRge(J), True)
        mToAdr = mToAdr & pRow + 1 & ":" & mToAdr & pRow + pNRow - 1
        .Range(mFmAdr).Copy
        If pCopyOnly_Val_Fmt Then
            .Range(mToAdr).PasteSpecial xlPasteFormulas
            .Range(mToAdr).PasteSpecial xlPasteFormats
        Else
            .Range(mToAdr).PasteSpecial xlPasteAll
        End If
    Next
End With
End Function

Function WsCpyVal(oWs As Worksheet, pFmWs As Worksheet, pToWsNm$) As Boolean
Const cSub$ = "WsCpyVal"
If Add_Ws(oWs, pFmWs.Parent, pToWsNm) Then ss.A 1: GoTo E
pFmWs.Cells.Copy
oWs.Select
oWs.Application.Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
oWs.Application.Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Exit Function
R: ss.R
E:
End Function

Sub WsDlt(A As Worksheet)
Dim mXls As Excel.Application: Set mXls = A.Application
mXls.DisplayAlerts = False
A.Delete
mXls.DisplayAlerts = True
Exit Sub
R: ss.R
E:
End Sub

Function WsFrzCell(A As Worksheet) As Range
Dim oCno, oRno
'Aim: Find the {oFreezedCellAdr} of {A}
A.Activate
Dim mWb As Workbook: Set mWb = A.Parent
Dim mWin As Window: Set mWin = A.Application.ActiveWindow
If mWin.Panes.Count <> 4 Then MsgBox "Given worksheet [" & A.Name & "] does not have 4 panes to find the Freezed Cell", vbCritical:  Exit Function
Dim mA$: mA = mWin.Panes(1).VisibleRange.Address
Dim mP%: mP = InStr(mA, ":")
mA = Mid(mA, mP + 1)
Dim mRge As Range: Set mRge = A.Range(mA)
Set WsFrzCell = mRge
End Function

Sub WsFrzCell__Tst()
'Debug.Print Application.Workbooks.Count
'Debug.Print Application.Workbooks(1).FullName
'Dim mWs As Worksheet: Set mWs = Application.Workbooks(1).Sheets("Input - HKDP")
'Dim mRno&, mCno As Byte: If Fnd_FreezedCell(mRno, mCno, mWs) Then Stop
'Debug.Print mRno & CtComma & mCno
'Dim mSqLeft As cSq, mSqTop As cSq
End Sub

Sub WsInsCell(A As Worksheet, pLoCol$, pRnoBeg&, pNRow&)
Dim mAyCol$(), J As Byte
mAyCol = Split(pLoCol, CtComma)
With A
    For J = 0 To UBound(mAyCol)
        .Range(mAyCol(J) & pRnoBeg & ":" & mAyCol(J) & pRnoBeg + pNRow - 1).Insert xlShiftToRight
    Next
End With
End Sub

Function WsNew(Optional WsNm$) As Worksheet
Dim Wb As Workbook: Set Wb = WbNew
WbKeepFstWs Wb
Dim O As Worksheet
Set O = WbFstWs(Wb)
If WsNm <> "" Then O.Name = WsNm
Set WsNew = O
End Function

Sub WsNewXNm(A As Worksheet, XNm$, ReferToAdr$)
A.Names.Add XNm, ReferToAdr$
End Sub

Sub WsReplInWb__Tst()
Dim mFx$: mFx = "c:\tmp\aa.xls"
Dim mWs As Worksheet, mWb As Workbook
If Crt_Wb(mWb, mFx, True, "Sheet1") Then GoTo E
If Add_Ws_ByLnWs(mWb, "Sheet2,Sheet3,Sheet4") Then GoTo E
Set mWs = mWb.Sheets("Sheet1"): If Set_Ws_ByLpAp(mWs, 1, "Sheet#,bbb,xxx", "Sheet1", 11111, 111119) Then GoTo E
Set mWs = mWb.Sheets("Sheet2"): If Set_Ws_ByLpAp(mWs, 1, "Sheet#,bbb,xxx", "Sheet2", 22222, 222229) Then GoTo E
Set mWs = mWb.Sheets("Sheet3"): If Set_Ws_ByLpAp(mWs, 1, "Sheet#,bbb,xxx", "Sheet3", 33333, 333339) Then GoTo E
Set mWs = mWb.Sheets("Sheet4"): If Set_Ws_ByLpAp(mWs, 1, "Sheet#,bbb,xxx", "Sheet4", 44444, 444449) Then GoTo E
mWb.Application.Visible = True
MsgBox "Sheet2 will be replaced by Sheet4 and Sheet2 will be deleted", , "Repl_Ws"
If Repl_Ws_InWb(mWb, "Sheet2", "Sheet4") Then GoTo E
Exit Sub
E:
End Sub

Sub WsRmvQt(A As Worksheet)
While A.QueryTables.Count > 0
    A.QueryTables(1).Delete
Wend
End Sub

Sub WsRpl(pWbTar As Workbook, pWbSrc As Workbook, pWsNm$)
'Aim: replace the {pWs} in {pWbTar} by {WbSrc}
Const cSub$ = "Repl_Ws_In2Wb"
On Error GoTo R
Dim mWsSrc As Worksheet: If Fnd_Ws(mWsSrc, pWbSrc, pWsNm) Then ss.A 1: GoTo E
Dim mWsTar As Worksheet: If Fnd_Ws(mWsTar, pWbTar, pWsNm) Then If Add_Ws(mWsTar, pWbTar, pWsNm) Then ss.A 2: GoTo E
If Repl_Ws(mWsTar, mWsSrc) Then ss.A 3: GoTo E
Exit Sub
R: ss.R
E:
End Sub

Sub WsRpl__Tst()
Const cFx1$ = "c:\tmp\aa.xls"
Const cFx2$ = "c:\tmp\bb.xls"
Dim mWb1 As Workbook: If Crt_Wb(mWb1, cFx1, True) Then Stop: GoTo E
Dim mWb2 As Workbook: If Crt_Wb(mWb2, cFx2, True) Then Stop: GoTo E
If Cls_Wb(mWb2, True) Then Stop: GoTo E
If Opn_Wb(mWb2, cFx2, True) Then Stop: GoTo E

Dim mWs1 As Worksheet: Set mWs1 = mWb2.Sheets(1)
Dim mWs2 As Worksheet: Set mWs2 = mWb2.Sheets(1)
mWs2.Range("A1").Value = "From"
If Repl_Ws_In2Wb(mWb1, mWb2, "ToBeDelete") Then Stop: GoTo E
mWb1.Application.Visible = True
Stop
GoTo X
E:
X: Cls_Wb mWb1
   Cls_Wb mWb2
End Sub

Sub WsRpl1(pWsTar As Worksheet, pWsSrc As Worksheet)
'Aim: replace the {pWsTar} by {pWsSrc} and delete {pWsTar}.  The 2 worksheets may in different wb.
'     If the workbook holding pWsSrc has only one worksheet, add a new ws will be added.
'     The pWsTar name will be preverse.
Const cSub$ = "Repl_Ws"
On Error GoTo R
Dim mWb As Workbook: Set mWb = pWsSrc.Parent
If mWb.Sheets.Count = 1 Then mWb.Sheets.Add
Dim mWsNm$: mWsNm = pWsTar.Name
pWsTar.Name = Format(Now, "yyyymmdd hhmmss")
pWsSrc.Move After:=pWsTar
If Dlt_Ws(pWsTar) Then ss.A 1: GoTo E
Exit Sub
R: ss.R
E:
End Sub

Sub WsRplInFx(Fx$, pWsNmTar$, pWsNmSrc$)
'Aim: replace the {pWsNmTar} by {pWsNmSrc} in same {Fx} and delete {pWsNmSrc}
Const cSub$ = "Repl_Ws_InFx"
Dim mWb As Workbook: If Opn_Wb_RW(mWb, Fx) Then ss.A 1: GoTo E
If Repl_Ws_InWb(mWb, pWsNmTar, pWsNmSrc) Then ss.A 2: GoTo E
If Cls_Wb(mWb, True) Then ss.A 3: GoTo E
Exit Sub
R: ss.R
E:
End Sub

Sub WsRplInFx__Tst()
Dim mFx$: mFx = "c:\tmp\aa.xls"
Dim mWs As Worksheet, mWb As Workbook
If Crt_Wb(mWb, mFx, True, "Sheet1") Then Stop
If Add_Ws_ByLnWs(mWb, "Sheet2,Sheet3,Sheet4") Then Stop
Set mWs = mWb.Sheets("Sheet1"): If Set_Ws_ByLpAp(mWs, 1, "Sheet#,bbb,xxx", "Sheet1", 11111, 111119) Then Stop
Set mWs = mWb.Sheets("Sheet2"): If Set_Ws_ByLpAp(mWs, 1, "Sheet#,bbb,xxx", "Sheet2", 22222, 222229) Then Stop
Set mWs = mWb.Sheets("Sheet3"): If Set_Ws_ByLpAp(mWs, 1, "Sheet#,bbb,xxx", "Sheet3", 33333, 333339) Then Stop
Set mWs = mWb.Sheets("Sheet4"): If Set_Ws_ByLpAp(mWs, 1, "Sheet#,bbb,xxx", "Sheet4", 44444, 444449) Then Stop
If Cls_Wb(mWb, True) Then Stop
MsgBox "Sheet2 will be replaced by Sheet4 and Sheet2 will be deleted", , "Repl_Ws"
If Repl_Ws_InFx(mFx, "Sheet2", "Sheet4") Then Stop
If Opn_Wb_R(mWb, mFx) Then Stop
mWb.Application.Visible = True
End Sub

Sub WsRplInWb(pWb As Workbook, pWsNmTar$, pWsNmSrc$)
'Aim: replace the {pWsNmTar$} by {pWsNmTar} in {pWb} and delete {pWsNmTar}
Const cSub$ = "Repl_Ws_InWb"
On Error GoTo R
Dim mWsTar As Worksheet: If Fnd_Ws(mWsTar, pWb, pWsNmTar) Then ss.A 1: GoTo E
Dim mWsSrc As Worksheet: If Fnd_Ws(mWsSrc, pWb, pWsNmSrc) Then ss.A 2: GoTo E
If Repl_Ws(mWsTar, mWsSrc) Then ss.A 3: GoTo E
Exit Sub
R: ss.R
E:
End Sub

Sub WsSetChtObj(A As Worksheet, pAyK$(), pAyV$())
Dim I As ChartObject
For Each I In A.ChartObjects
    ChtTitSet I.Chart.ChartTitle, pAyK, pAyV
Next
End Sub

Sub WsSetFze(A As Worksheet, Adr$)
With A.Range(Adr)
    .Activate
    .Select
End With
ActiveWindow.FreezePanes = True
End Sub

Sub WsSetOLEObjPrm(A As Excel.Worksheet, Fx$, CtlCnt%, PrpNm$, V)
'Aim: Assume there are CtlCnt comboxbox control object in the Ws-{A} with name XXX01, ... XXXnn, where XXX is {Pfx}, nn is 1..CtlCnt
'     It is required to set the property pPrp for each of the control by the value V
Dim J%, N$, O As OLEObject
For J = 1 To CtlCnt
    Set O = A.OLEObjects(Fx & Format(J, "00"))
    Select Case PrpNm
    Case "ListFillRange":    O.ListFillRange = V
    Case "PrintObject":      O.PrintObject = V
    Case "Height":           O.Height = V
    Case "Height":           O.Height = V
    Case "ListRows":         O.ListRows = V
    Case Else
    Stop
    End Select
Next
End Sub

Sub WsSetZoom(A As Worksheet, Zoom%)
Dim Wb As Workbook: Set Wb = A.Parent
A.Activate
Dim W As Window
For Each W In Wb.Windows
    W.Zoom = Zoom
Next
End Sub

Sub WsShwAllCols(A As Worksheet)
A.Application.ScreenUpdating = False
Dim OutLin As Outline: Set OutLin = A.Outline
Dim J As Byte
For J = 0 To 8
    OutLin.ShowLevels , J
Next
A.Application.ScreenUpdating = True
End Sub

Sub WsShwAllDta(A As Worksheet)
WsShwAllCols A
On Error Resume Next
A.ShowAllData
End Sub

Function WsTwoSq(pWs As Worksheet) As Variant()
'Aim: Find 2Sq: {oSqLeft} & {oSqTop} by {pWs}, {pRno} & {pCno}.  {pRno} & {pCno} are bottom right corner of pane of the freezed window.
Dim oSqLeft, oSqTop
Const cSub$ = "Fnd_TwoSq"
'If TypeName(oSqLeft) = "Nothing" Then Set oSqLeft = New cSq
'If TypeName(oSqTop) = "Nothing" Then Set oSqTop = New cSq
'Find the Freeze Cell
Dim mFreezeRno&, mFreezeCno As Byte: If Fnd_FreezedCell(mFreezeRno, mFreezeCno, pWs) Then ss.A 1: GoTo E
'Detect mLasRno&, mLasCno
Dim mLasRno&, mLasCno As Byte
Dim mRge As Range: Set mRge = pWs.Cells.SpecialCells(xlCellTypeLastCell)
mLasRno = mRge.Row
mLasCno = mRge.Column
'Work From LasRno to pRno+1 to find first non-empty row so that oSqLeft is find
Dim iRno&, iCno%, mIsEmpty As Boolean
For iRno = mLasRno + 1 To mFreezeRno + 1 Step -1
    mIsEmpty = True
    For iCno = 1 To mFreezeCno
        If Not IsEmpty(pWs.Cells(iRno, iCno).Value) Then mIsEmpty = False: Exit For
    Next
    If mIsEmpty Then Exit For
Next
With oSqLeft
    .Cno1 = 1
    .Cno2 = mFreezeCno
    .Rno1 = mFreezeRno + 1
    .Rno2 = iRno - 1
End With
'Work From LasCno to pCno+1 to find first non-empty column so that oSqTop is find
For iCno = mLasCno + 1 To mFreezeCno + 1 Step -1
    mIsEmpty = True
    For iRno = 1 To mFreezeRno
        If Not IsEmpty(pWs.Cells(iRno, iCno).Value) Then mIsEmpty = False: Exit For
    Next
    If mIsEmpty Then Exit For
Next
With oSqTop
    .Cno1 = mFreezeCno + 1
    .Cno2 = iCno - 1
    .Rno1 = 1
    .Rno2 = mFreezeRno
End With
Exit Function
E:
End Function

Sub WsVis(A As Worksheet)
A.Application.Visible = True
End Sub
