Attribute VB_Name = "nXls_nObj_nCell_nDo_Cell"
Option Compare Database
Option Explicit

Sub CellAct(Cell As Range)
RgWs(Cell).Activate
With RgRC(Cell, 1, 1)
    .Activate
    .Select
End With
End Sub

Sub CellCrtPt(Cell As Range, RgNam$, _
    RowFnStr$, ColFnStr$, DtaFnStr$, _
    Optional RowTotFnStr$ = "", Optional ColTotFnStr$ = "")

'Aim: Create a new Ws of name {pWsnPt} having a Pt from a data source as defined in name {RgNam}
'     pFLst is in format [<<FCaption>>:]<<FNam>>
Dim O As PivotTable:
    Dim PtNm$: PtNm = "Pt_" & RgWs(Cell).Name
    Set O = RgWb(Cell).PivotCaches.Add(xlDatabase, RgNam).CreatePivotTable("", PtNm)
    
O.PivotCache.MissingItemsLimit = xlMissingItemsNone
Dim Ay$(), J%, mNmF$, mFCaption$
With O
    'Set F col
    Ay = NmBrk(ColFnStr)
    Dim F As PivotField
    For J = UBound(Ay) To 0 Step -1
        If Brk_ColonAs_ToCaptionNm(mFCaption, mNmF, Ay(J)) Then ss.A 1:
        
        Set F = .PivotFields(mNmF)
        With F
            .Orientation = xlColumnField
            .Position = 1
            .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            If mFCaption <> mNmF Then .Caption = mFCaption
        End With
    Next

    'Set F Row
    Ay = NmBrk(RowFnStr)
    For J = UBound(Ay) To LBound(Ay) Step -1
        If Brk_ColonAs_ToCaptionNm(mFCaption, mNmF, Ay(J)) Then ss.A 2:
        
        Set F = .PivotFields(mNmF)
        With F
            .Orientation = xlRowField
            .Position = 1
            .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
            If mFCaption <> mNmF Then .Caption = mFCaption
        End With
    Next

    'Set F Data
    Ay = NmBrk(DtaFnStr)
    For J = UBound(Ay) To LBound(Ay) Step -1
        If Brk_Str_Both(mNmF, mFCaption, Ay(J), ":") Then ss.A 1:
        Set F = .PivotFields(mNmF)
        With F
            .Orientation = xlDataField
            .Position = 1
            .Function = xlSum
            If mFCaption <> "" Then .Caption = mFCaption
        End With
    Next

    'Set Data Fields as col
    With O.DataPivotField
        .Orientation = xlColumnField
        .Position = 1
    End With
End With
PtSrt O
End Sub

Sub CellDrpCmt(A As Range)
Dim C As Comment: Set C = A.Comment
If TypeName(C) = "Nothing" Then Exit Sub
C.Delete
End Sub

Sub CellInsRowAbove(Cell As Range)
Dim mObj As Object: Set mObj = Excel.Application.Selection
If TypeName(mObj) <> "Range" Then MsgBox "A row must be selected": Exit Sub
Dim mRge As Range: Set mRge = Excel.Application.Selection
If mRge.Rows.Count <> 1 Then MsgBox "Only a row can be selected": Exit Sub
Dim mWs As Worksheet: Set mWs = mRge.Parent
WsShwAllCols mWs
mRge.EntireRow.Insert xlDown
Dim mRowFm As Range: Set mRowFm = mWs.Rows(mRge.Row)
Dim mRowTo As Range: Set mRowTo = mWs.Rows(mRge.Row - 1)
mRowFm.Copy
mRowTo.PasteSpecial xlPasteAll
mRowTo.Range("B1").Value = Null
Excel.Application.CutCopyMode = False
End Sub

Sub CellPutApDown(Cell As Range, ParamArray Ap())
Dim Av(): Av = Ap
CellPutAv Cell, Av, IsDown:=True
End Sub

Sub CellPutApRight(Cell As Range, ParamArray Ap())
Dim Av(): Av = Ap
CellPutAv Cell, Av, IsDown:=False
End Sub

Sub CellPutAv(Cell As Range, Av(), IsDown As Boolean)
Dim Sq
If IsDown Then Sq = AySqV(Av) Else Sq = AySqH(Av)
SqPutCell Sq, Cell
End Sub

