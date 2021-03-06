Attribute VB_Name = "nXls_nObj_nWb_nInf_Wb"
Option Compare Database
Option Explicit

Sub WbActAllWsA1(A As Workbook)
Dim iWs As Worksheet
For Each iWs In A.Sheets
    If iWs.Visible Then
        iWs.Activate
        iWs.Range("A1").Activate
    End If
Next
A.Sheets(1).Activate
End Sub

Function WbAddCsv(A As Workbook, Fcsv$, Optional WsNm$, Optional KeepCsv As Boolean) As Worksheet
'Aim: Add a new ws to {A} as {WsNm} from {Fcsv}.  If {WsNm} is not given, use {Fcsv} as worksheet name
'Open {Fcsv} & Set as <FmWb>
Dim oWs As Worksheet
FfnAsstExt Fcsv, ".csv", "WbAddCsv"
Dim FmWb As Workbook
    Set FmWb = A.Application.Workbooks.Open(Fcsv)
Dim WsNm1$
    WsNm1 = NonBlank(WsNm, Fct.Nam_FilNam(Fcsv, False))
'Add a new Ws as {oWs} of name {WsNm}
Set oWs = WbAddWs(A, WsNm1)

'Copy from <FmWb.Sheet1> and Paste to <oWs>
FmWb.Sheets(1).Cells.Copy
oWs.Cells.PasteSpecial xlPasteAll
oWs.Activate
oWs.Range("A1").Select

'Close <FmWb>
WbCls FmWb, NoSav:=True

'Kill Fcsv
If Not KeepCsv Then FfnDlt Fcsv
Set WbAddCsv = oWs
End Function

Function WbAddWs(A As Workbook, WsNm$) As Worksheet
Set WbAddWs = WbAddWsAtEnd(A, WsNm)
End Function

Function WbAddWsAft(A As Workbook, WsNm$, Aft) As Worksheet
Dim O As Worksheet: Set O = A.Sheets.Add(, Aft)
O.Name = WsNm
Set WbAddWsAft = O
End Function

Function WbAddWsAtBeg(A As Workbook, WsNm$) As Worksheet
Dim O As Worksheet: Set O = A.Sheets.Add
O.Name = WsNm
Set WbAddWsAtBeg = O
End Function

Sub WbAddWsAtBeg__Tst()
Dim A As Workbook: Set A = WbNew
WbAddWsAtBeg A, "AAA"
Stop
End Sub

Function WbAddWsAtEnd(A As Workbook, WsNm$) As Worksheet
Dim Las As Worksheet: Set Las = WbLasWs(A)
Set WbAddWsAtEnd = WbAddWsAft(A, WsNm, Las)
End Function

Function WbAddWsBef(A As Workbook, WsNm$, Bef) As Worksheet
Dim O As Worksheet: Set O = A.Sheets.Add(Bef)
O.Name = WsNm
Set WbAddWsBef = O
End Function

Sub WbCls(W As Workbook, Optional NoSav As Boolean)
On Error GoTo X
XlsDspAlertPush W.Application, False
W.Close Not NoSav
XlsDspAlertPop
X:
End Sub

Sub WbClsNosav(A As Workbook)
WbCls A, NoSav:=True
End Sub

Function WbFstWs(A As Workbook) As Worksheet
Set WbFstWs = A.Sheets(1)
End Function

Function WbHasWs(A As Workbook, WsIdx) As Boolean
On Error GoTo R
Dim W As Worksheet: Set W = A.Sheets(WsIdx)
WbHasWs = True
Exit Function
R:
End Function

Sub WbHasWs__Tst()
Const WsNm$ = "xxxxx"
Dim Wb As Workbook
Set Wb = WbNew
Debug.Assert WbHasWs(Wb, WsNm) = False
WbAddWs Wb, WsNm
Debug.Assert WbHasWs(Wb, WsNm) = True
WbClsNosav Wb
End Sub

Function WbHasXNm(A As Workbook, XNm$) As Boolean
On Error GoTo R
Dim oNm As Name
Set oNm = A.Names(XNm)
WbHasXNm = True
Exit Function
R:
End Function

Sub WbKeepFstWs(A As Workbook)
Dim J%
For J = A.Sheets.Count To 2 Step -1
    WbWs(A, J).Delete
Next
End Sub

Function WbLasWs(A As Workbook) As Worksheet
Set WbLasWs = A.Sheets(A.Sheets.Count)
End Function

Function WbLy(A As Workbook) As String()
'Dim iWs As Worksheet, iQt As QueryTable, iPt As PivotTable
'For Each iWs In pWb.Worksheets
'    If iWs.PivotTables.Count > 0 Then
'        Prt_Ln pFno, Fct.UnderlineStr("Worksheet " & iWs.Name & " (PivotTables)", "-")
'        For Each iPt In iWs.PivotTables
'            Prt_Ln pFno, PtToStr(iPt)
'        Next
'        Prt_Ln pFno
'    End If
'    If iWs.QueryTables.Count > 0 Then
'        Prt_Ln pFno, Fct.UnderlineStr("Worksheet " & iWs.Name & " (QueryTables)", "-")
'        For Each iQt In iWs.QueryTables
'            Prt_Ln pFno, QtToStr(iQt)
'        Next
'        Prt_Ln pFno
'    End If
'Next
End Function

Function WbNew(Optional Fx$, Optional Vis As Boolean) As Workbook
Dim O As Workbook: Set O = Appx.Workbooks.Add
Appx.Visible = Vis
If Fx <> "" Then WbSavAs O, Fx
Set WbNew = O
End Function

Sub WbRmvAllWsExcept__Tst()
Dim W As Workbook
Set W = WbNew
WbAddWs W, "Sheet2"
WbAddWs W, "Sheet3"
WbAddWs W, "Sheet4"
WbRmvAllWsExcp W, "Sheet2"
WbVis W
Stop
WbClsNosav W
End Sub

Sub WbRmvAllWsExcp(A As Workbook, ExcpNmstr$)
Dim N$(): N = WbWsNy(A)
Dim OWsNy$(): OWsNy = NmstrExcp(ExcpNmstr, N)
WbRmvWsNy A, OWsNy
End Sub

Sub WbRmvWs(A As Workbook, WsNm)
WbWs(A, WsNm).Delete
End Sub

Sub WbRmvWsNy(A As Workbook, WsNy$())
If AyIsEmpty(WsNy) Then Exit Sub
Dim I
For Each I In WsNy
    WbRmvWs A, I
Next
End Sub

Sub WbSav(W As Workbook)
XlsDspAlertPush W.Application
W.Save
XlsDspAlertPop
End Sub

Sub WbSavAs(W As Workbook, Fx$, Optional FilFmt As XlFileFormat = XlFileFormat.xlWorkbookDefault)
XlsDspAlertPush W.Application, False
W.SaveAs Fx, FilFmt
XlsDspAlertPop
End Sub

Sub WbSetMin(pWb As Workbook)
Dim iWin As Window
For Each iWin In pWb.Windows
    If iWin.WindowState <> xlMinimized Then iWin.WindowState = xlMinimized
Next
End Sub

Sub WbSetPjNm(A As Workbook)
If FfnExt(A.Name) <> ".xlam" Then Er "SetWbPjNm Err: Given {Wb} name must have extension [.xlam]", A.FullName
Dim Nm$: Nm = FfnFnn(A.Name)
Dim I As vbproject
For Each I In A.Application.Vbe.VBProjects
    If I.FileName = A.FullName Then
        I.Name = Nm
        WbSav A
        Exit Sub
    End If
Next
End Sub

Sub WbSetPjNm__Tst()
Dim F$: F = TmpFil(".xlam")
Dim W As Workbook: Set W = WbNew(F)
WbSetPjNm W
End Sub

Sub WbShwLvl1(A As Workbook)
Dim W As Worksheet
For Each W In A.Sheets
    W.Outline.ShowLevels 1, 1
Next
End Sub

Function WbToStr$(A As Workbook)
On Error GoTo R
WbToStr = A.FullName
Exit Function
R: WbToStr = "WbToStr error: Msg=" & Err.Description
End Function

Sub WbVis(A As Workbook)
A.Application.Visible = True
End Sub

Function WbWs(W As Workbook, WsIdx) As Worksheet
Set WbWs = W.Sheets(WsIdx)
End Function

Function WbWsAy(A As Workbook, Optional InclHid) As Worksheet()
Dim O() As Worksheet
WbWsAy = CollOy(A.Sheets, O)
End Function

Function WbWsNy(A As Workbook, Optional InclHid As Boolean) As String()
WbWsNy = OyPrp_Nm(WbWsAy(A, InclHid))
End Function

Sub WbWsNy__Tst()
Dim Wb As Workbook
Set Wb = Appx.Workbooks.Add
AyBrw WbWsNy(Wb)
Stop
WbClsNosav Wb
End Sub

Sub WsCls(A As Worksheet, Optional NoSav As Boolean)
WbCls WsWb(A), NoSav
End Sub

Sub WsClsNoSav(A As Worksheet)
WbClsNosav WsWb(A)
End Sub
