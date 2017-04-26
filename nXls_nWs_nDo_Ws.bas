Attribute VB_Name = "nXls_nWs_nDo_Ws"
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
'Aim: Add the content of {pFxFm} at the end of {pWsTo} provided that they are same layout
'Assume: [1 Ws in pFxFm] & [Same Layout]
''[1 Ws in pFxFm] There is only one ws in {pFxFm} having the same name of the file name of {pFxFm}
''[Same Layout]       The column headings of {pFxFm} & {pWsTo} should be the same
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
Dim mNewWs As Worksheet: If Add_Ws(mNewWs, mWbTo, mWsNmTo) Then ss.A 2: GoTo E
mNewWs.Cells.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
mNewWs.Cells.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Set oWsTar = mNewWs
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
E: WsCpyRow = True: ss.B cSub, cMod, "pWs,pRnoFm,pRnoTo", ToStr_Ws(pWs), pRnoFm, pRnoTo
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

Function WsCpyVal(OWs As Worksheet, pFmWs As Worksheet, pToWsNm$) As Boolean
Const cSub$ = "WsCpyVal"
If Add_Ws(OWs, pFmWs.Parent, pToWsNm) Then ss.A 1: GoTo E
pFmWs.Cells.Copy
OWs.Select
OWs.Application.Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
OWs.Application.Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Exit Function
R: ss.R
E: WsCpyVal = True: ss.B cSub, cMod, "pFmWs,pToWsNm", ToStr_Ws(pFmWs), pToWsNm
End Function

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

Sub WsRmvQt(A As Worksheet)
While A.QueryTables.Count > 0
    A.QueryTables(1).Delete
Wend
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
