Attribute VB_Name = "nXls_nDta_Dt"
Option Compare Database
Option Explicit

Function DtNewWs(A As Dt) As Worksheet
Dim O As Worksheet
Set O = WsNew
DtPutCell A, WsA1(O)
Set DtNewWs = O
End Function

Sub DtNewWs__Tst()
Dim Fny$()
Dim DrAy()
    Fny = Split("AA BBD", " ")
    DrAy = Array( _
        Array("'001", "A"), _
        Array("'002", "B"))

Dim Dt As Dt
    Dt = DtNew(Fny, DrAy)
WsVis DtNewWs(Dt)
End Sub

Function DtPKey(A As Dt, Optional PkFnStr$) As Dictionary
Dim PkFny$()
    If PkFnStr = "" Then
        PkFny = ApSy(A.Fny(0))
    Else
        PkFny = NmstrBrk(PkFnStr)
    End If
Dim PkIdx&()
    PkIdx = AyIdxAy(A.Fny, PkFny)
Dim O As New Dictionary
    Dim J&
    Dim Dr, DrAy(), Pk$
    DrAy = A.DrAy
    For J = 0 To DtURec(A)
        Dr = DrAy(J)
        Pk = Join(DrSel(Dr, PkIdx), "|")
        If O.Exists(Pk) Then
            Er "{Tn} of {Row} has dup {Pk} of {Pk-Fny}", A.Tn, J, Pk, FnyToStr(PkFny)
        End If
        O.Add Pk, J
    Next
Set DtPKey = O
End Function

Sub DtPKey__Tst()
Dim O As Dictionary
AyBrw DtScLy(DicDt(DtPKey(TblDt("Permit"), "PermitNo")))
'Set O = DtPKey(TblDt("Permit"), "PermitNo")
'DicBrw O

End Sub

Function DtPutCell(A As Dt, Cell As Range, Optional NoListObj As Boolean) As Range
Dim O As Range
Dim Sq: Sq = DtSq(A)
Set O = CellReSz(Cell, Sq)
O.Value = Sq
Set DtPutCell = O
If Not NoListObj Then ListObjCrt O
End Function

Sub DtPutCell__Tst()
Dim Ws As Worksheet: Set Ws = WsNew
Dim Cell As Range: Set Cell = WsRC(Ws, 2, 2)
Dim Rg As Range: Set Rg = DtPutCell(DtSample1, Cell)

End Sub

Function DtSq(A As Dt)
Dim UC&: UC = DtColUB(A)
Dim UR&: UR = UB(A.DrAy)
Dim IR&, IC&, Dr, DrAy(), R&
Dim O(): ReDim O(1 To UR + 2, 1 To UC + 1)
DrAy = A.DrAy
For IC = 0 To UB(A.Fny)
    O(1, IC + 1) = A.Fny(IC)
Next
For IR = 0 To UR
    Dr = DrAy(IR)
    R = IR + 2
    For IC = 0 To UB(Dr)
        O(R, IC + 1) = Dr(IC)
    Next
Next
DtSq = O
End Function

Function DtWs(Dt As Dt, Optional WsNm$ = "Data") As Worksheet
Dim O As Worksheet
Set O = WsNew(WsNm)
ListObjCrt DtPutCell(Dt, WsA1(O))
Set DtWs = O
End Function

