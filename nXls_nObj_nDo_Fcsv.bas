Attribute VB_Name = "nXls_nObj_nDo_Fcsv"
Option Compare Database
Option Explicit
Const C_Mod$ = "nXls_nImp_Fcsv"

Sub aa()
CsvPthPt__Tst
End Sub

Function CsvPthLo(CsvPth$, Sql$, Cell As Range, Optional LoNm$) As ListObject
Dim Ws As Worksheet: Set Ws = RgWs(Cell)
Dim LoNm1$: LoNm1 = LoNmNz(LoNm, Ws)
Dim Src$
Dim Pth$: Pth = CsvPth
Src = WcCnnStrCsvPth(Pth)
Dim LO As ListObject
Set LO = Ws.ListObjects.Add(SourceType:=0, Source:=Src, Destination:=Cell)
With LO.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Sql
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = LoNmNz(LoNm, Ws)
        .Refresh BackgroundQuery:=False
End With
Set FcsvLo = LO
End Function

Sub CsvPthLo__Tst()
Dim Ws As Worksheet: Set Ws = WsNew
Dim Fcsv$: Fcsv = "F - Copy (3).csv"
Dim Pth$: Pth = "N:\SAPACCESSREPORTS\DUTYPREPAY5\TSTRES\NXLS_WCLNK\"
Dim Sql$: Sql = "Select * from [F1.csv] a,[F2.csv] b"
CsvPthLo Pth, Sql, Ws.Range("B4")
WsVis Ws
Stop
WsClsNoSav Ws
End Sub

Function CsvPthPc(CsvPth$, Sql$, Wb As Workbook, Optional WcNm$, Optional Des$) As PivotCache
Dim Wc As WorkbookConnection: Set Wc = WcLnkCsvPth(Wb, CsvPth, Sql, WcNm, Des)
Set CsvPthPc = Wb.PivotCaches.Create(SourceType:=xlExternal, SourceData:=Wc, Version:=xlPivotTableVersion14)
End Function

Sub CsvPthPt(CsvPth$, Sql$, Cell As Range, PtFmtrLy$())
If CsvPth = "" Then Err.Raise 1, "CsvPthPt: CsvPth cannot be blank"
If Sql$ = "" Then Err.Raise 1, "CsvPthPt: Sql cannot be blank"
Dim Wb As Workbook: Set Wb = RgWb(Cell)
Dim Pc As PivotCache: Set Pc = CsvPthPc(CsvPth, Sql, Wb)
Dim Pt As PivotTable: Set Pt = PcPt(Pc, Cell)
If AyIsEmpty(PtFmtrLy) Then Exit Sub
PtFmt Pt, PtFmtr(PtFmtrLy)
End Sub

Sub CsvPthPt__Tst()
Dim CsvPth$
Dim Sql$
Dim Cell As Range
Dim A$()
    CsvPth = "N:\SAPACCESSREPORTS\DUTYPREPAY5\TSTRES\NXLS_WCLNK\"
    Dim Ws As Worksheet: Set Ws = WsNew
    WsVis Ws
    Set Cell = WsA1(Ws)
    Sql = "Select * from [F1.csv]"

Dim F$()
Push F, "Fny: aa bb CC e DD"
Push F, "Row: aa bb CC"
Push F, "Pag: e"
Push F, "Dta: DD"
Push F, "GrandColTot: True 40"
Push F, "GrandRowTot: False"
Push F, "SubTot: AA BB"
Push F, "OpnInd: False"
Push F, "Wdt: 40: AA CC"
Push F, "OutLin: 2: BB"
Push F, "OutLin: 3: CC"
Push F, "Lbl: AA : AA-Lbl"
Push F, "Lbl: DD : DD-Lbl "
Push F, "DtaSum: DD Sum #,##0.0"
Dim Fmtr As PtFmtr
CsvPthPt CsvPth, Sql, Cell, F
End Sub

Function FcsvLo(Fcsv$, Cell As Range, Optional LoNm$) As ListObject
Dim Sql$
Dim Pth$
Dim Fn$
Fn = FfnFn(Fcsv)
Pth = FfnPth(Fcsv)
Sql = FmtQQ("Select * from [?]", Fn)
Set FcsvLo = FcsvLo(Pth, Sql, Cell, LoNm)
End Function

Sub FcsvLo__Tst()
Dim Fcsv$: Fcsv = "N:\SAPACCESSREPORTS\DUTYPREPAY5\TSTRES\NXLS_WCLNK\F1.csv"
Dim Ws As Worksheet: Set Ws = WsNew
Dim Cell As Range: Set Cell = WsA1(Ws)
FcsvLo Fcsv, Cell
WsVis Ws
Stop
WsClsNoSav Ws
End Sub

Function PcPt(A As PivotCache, Cell As Range, Optional PtNm$) As PivotTable
Dim Ws As Worksheet: Set Ws = RgWs(Cell)
Dim B As PivotTables: Set B = Ws.PivotTables
Dim N$: N = PtNmNz(PtNm, Ws)
Set PcPt = B.Add(A, Cell, N)
End Function

Function PtFny(A As PivotTable) As String()
Dim F As PivotField
Dim O$()
For Each F In A.PivotFields
    Push O, F.Name
Next
PtFny = O
End Function

Function PtNmIsExist(PtNm$, A As Worksheet) As Boolean
On Error GoTo X
PtNmIsExist = A.PivotTables(PtNm).Name = PtNm
Exit Function
X:
End Function

Function WcCnnStrCsvPth$(CsvPth$)
WcCnnStrCsvPth = WcCnnStrCsv(CsvPth)
End Function

Function WcCnnStrFb$(Fb$, T$)

End Function

Function WcCnnStrFcsv$(Fcsv$)
WcCnnStrFcsv = WcCnnStrCsv(FfnPth(Fcsv))
End Function

Sub WcCnnStrFcsv__Tst()
Dim A$: A = WcCnnStrFcsv(ZTstResFcsv)
Debug.Assert A = 1
StrBrw A
End Sub

Function WcCnnStrFx$(Fx$, WsNm$)
End Function

Function WcLnk(A As Workbook, WcNm$, Des$, CnnStr$, CmdTxt$, CmdTy As XlCmdType) As WorkbookConnection
Dim N$: N = WcNmNz(WcNm, A)
Set WcLnk = A.Connections.Add(N, Des, CnnStr, CmdTxt, CmdTy)
End Function

Function WcLnkCsvPth(A As Workbook, CsvPth$, Sql$, Optional WcNm$, Optional Des$) As WorkbookConnection
Dim C$: C = WcCnnStrCsvPth(CsvPth)
Set WcLnkCsvPth = WcLnk(A, WcNm, Des, C, Sql, xlCmdSql)
End Function

Function WcLnkFb(A As Workbook, Fb$, T$, Optional WcNm$, Optional Des$) As WorkbookConnection
Dim C$: C = WcCnnStrFb(Fb, T)
Set WcLnkFb = WcLnk(A, WcNm, Des, C, T, xlCmdTable)
End Function

Function WcLnkFcsv(A As Workbook, Fcsv$, Optional WcNm$, Optional Des$) As WorkbookConnection
Dim C$: C = WcCnnStrFcsv(Fcsv)
Dim T$: T = FfnFn(Fcsv)
Set WcLnkFcsv = WcLnk(A, WcNm, Des, C, Fcsv, xlCmdTable)
End Function

Sub WcLnkFcsv__Tst()
Dim W As Workbook
Dim Des$
Dim WcNm$
Dim T
Set W = WbNew
AppxShw
WcNm = WcNmNz("", W)
WcLnkFcsv W, ZTstResFcsv, WcNm
End Sub

Function WcLnkFx(A As Workbook, Fx$, WsNm$, Optional WcNm$, Optional Des$) As WorkbookConnection
Dim C$: C = WcCnnStrFx(Fx, WsNm)
Dim T$
Set WcLnkFx = WcLnk(A, WcNm, Des, C, T, xlCmdTable)
End Function

Function WcNmIsExist(WcNm, A As Workbook) As Boolean
On Error GoTo X
WcNmIsExist = A.Connections(WcNm).Name = WcNm
Exit Function
X:
End Function

Function WcNmNz$(WcNm$, A As Workbook)
If Not StrIsBlank(WcNm) Then WcNmNz = WcNm: Exit Function
Dim O$, J%
For J = 1 To 1000
    O = "WbCnn" & J
    If Not WcNmIsExist(O, A) Then WcNmNz = O: Exit Function
Next
Er "WcNmNz: Impossible!!"
End Function

Sub ZTstResFcsv__Tst()
Debug.Assert FfnIsExist(ZTstResFcsv) = True
End Sub

Private Function WcCnnStrCsv$(CsvPth$)
Const C$ = "OLEDB;Provider=MSDASQL.1;Extended Properties=""DBQ=?;" & _
        "Driver={Microsoft Access Text Driver (*.txt, *.csv)};Extensions=csv;MaxScanRow=1;FIL=text;MaxBufferSize=2048;"""
Dim O$: O = FmtQQ(C, CsvPth)
WcCnnStrCsv = O
End Function

Private Sub ZTstResBrw()
PthBrw TstResMdPth(C_Mod)
End Sub

Private Function ZTstResFcsv$(Optional No%)
ZTstResFcsv = TstResMdFcsv(C_Mod, No%)
End Function

Private Sub ZTstResFcsvBrw()
FtBrw ZTstResFcsv()
End Sub

