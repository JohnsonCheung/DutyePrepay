Attribute VB_Name = "nXls_nObj_nWs_nInf_Ws"
Option Compare Database
Option Explicit

Function WsA1(A As Worksheet) As Range
Set WsA1 = A.Range("A1")
End Function

Function WsCnoAyColAy(CnoAy%()) As String()
Dim O%()
WsCnoAyColAy = AyMapInto(CnoAy, O, "WsCnoCol")
End Function

Function WsCnoCol$(Cno%)
Dim A As Byte:      A = Asc("A")
Dim mA1 As Byte:    mA1 = (Cno - 1) \ 26
Dim mA2 As Byte:    mA2 = (Cno - 1) Mod 26
If mA1 = 0 Then WsCnoCol = Chr(A + mA2): Exit Function
WsCnoCol = Chr(A - 1 + mA1) & Chr(A + mA2)
End Function

Function WsColCno%(Col$)
Dim mC1$
Dim mC2$
If Len(Col) = 1 Then
    mC1 = UCase(Col)
    If "A" > mC1 Or mC1 > "Z" Then Exit Function
    WsColCno = Asc(mC1) - 64
    Exit Function
End If
If Len(Col) = 2 Then
    mC1 = UCase(Left(Col, 1))
    mC2 = UCase(Right(Col, 1))
    If "A" > mC1 Or mC1 > "Z" Then Exit Function
    If "A" > mC2 Or mC2 > "Z" Then Exit Function
    WsColCno = 26 * (Asc(mC1) - 64) + Asc(mC2) - 64
End If
End Function

Function WsColNxt$(Col$, NCol%)
WsColNxt = WsCnoCol(WsColCno(Col) + NCol)
End Function

Function WsCRR(A As Worksheet, C, R1, R2) As Range
Set WsCRR = A.Range(A.Cells(R1, C), A.Cells(R2, C))
End Function

Function WsHasXNm(A As Excel.Workbook, XNm$) As Boolean
On Error GoTo R
Dim oNm As Excel.Name
Set oNm = A.Names(XNm)
WsHasXNm = True
Exit Function
R:
End Function

Function WsIsInFx(WsNm$, Fx$) As Boolean
Dim W As Workbook
    Set W = FxWb(Fx)
WsIsInFx = WsIsInWb(WsNm, W)
WbCls W, NoSav:=True
End Function

Function WsIsInWb(WsNm$, Wb As Workbook) As Boolean
Dim W As Worksheet
For Each W In Wb.Sheets
    If W.Name = WsNm Then WsIsInWb = True: Exit Function
Next
End Function

Function WsLasCell(A As Worksheet) As Range
Set WsLasCell = A.Cells.SpecialCells(xlCellTypeLastCell)
End Function

Function WsMaxCno&(A As Worksheet)
WsMaxCno = A.Columns.Count
End Function

Function WsMaxRno&(A As Worksheet)
WsMaxRno = A.Rows.Count
End Function

Function WsHasLoNm(A As Workbook, LoNm$) As Boolean
On Error GoTo X
WsHasLoNm = A.ListObjects(LoNm).Name = LoNm
X:
End Function

Function WsQt(A As Worksheet, QtNm$) As QueryTable
Set WsQt = A.QueryTables(QtNm)
End Function

Function WsQtAy(A As Worksheet) As QueryTable()

End Function

Function WsRC(A As Worksheet, R, C) As Range
Set WsRC = A.Cells(R, C)
End Function

Function WsRCC(A As Worksheet, R, C1, C2) As Range
Set WsRCC = A.Range(A.Cells(R, C1), A.Cells(R, C2))
End Function

Function WsRCRC(A As Worksheet, R1, C1, R2, C2) As Range
Set WsRCRC = A.Range(WsRC(A, R1, C1), WsRC(A, R2, C2))
End Function

Function WsSq(A As Worksheet)
Dim LasCell As Range
    Set LasCell = WsLasCell(A)
Dim C&, R&
    C = LasCell.Column
    R = LasCell.Row
WsSq = WsRCRC(A, 1, 1, R, C).Value
End Function

Function WsWb(A As Worksheet) As Workbook
Set WsWb = A.Parent
End Function
