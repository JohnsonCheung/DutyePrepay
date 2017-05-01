Attribute VB_Name = "nXls_nObj_nRg_nInf_Rg"
Option Compare Database
Option Explicit

Function RgA1(A As Range) As Range
Set RgA1 = A(1, 1)
End Function

Function RgC(A As Range, C) As Range
Set RgC = RgRCRC(A, 1, C, RgNRow(A), C)
End Function

Function RgC1&(A As Range)
RgC1 = A.Column
End Function

Function RgC2&(A As Range)
RgC2 = A.Column + A.Columns.Count - 1
End Function

Function RgDicH(A As Range) As Dictionary
Set RgDicH = SqDic(SqTranspose(A.Value))
End Function

Function RgDicV(A As Range) As Dictionary
Set RgDicV = SqDic(A.Value)
End Function

Function RgEntireC(A As Range, C) As Range
Set RgEntireC = RgC(A, C).EntireColumn
End Function

Function RgEntireR(A As Range, R) As Range
Set RgEntireR = RgR(A, R).EntireRow
End Function

Function RgGpCnoAy(A As Range, Optional ExclSingleEleGp As Boolean) As Variant()
'Aim: Return FmToCnoDrAy().  Each element is an array of 2 values: FmCno & ToCno
'     The R1 of FmCno & ToCno of Range-A will have same cell-value
'     Note: empty cell block will not included.
If A.Count = 1 Then Exit Function
Stop
End Function

Function RgHasCmt(Rg As Range) As Boolean
On Error GoTo R
RgHasCmt = TypeName(Rg.Comment) <> "Nothing"
Exit Function
R:
End Function

Function RgKeyAdrDicH(A As Range) As Dictionary
'Aim: Find the KeyCnoDic {A}-Horizontal-Row.  Put each non-blank-non-dup-string to {ODic} with the cell-value as key and the {Cno} as value
Dim O As New Dictionary
    Dim LasCno&: LasCno = RgLasCnoFmEdge(A)
    Dim iCno&
    For iCno = A.Column To LasCno - A.Column + 1
        Dim V: V = A(0, iCno).Value
        If VarType(V) = vbString Then
            If O.Contains(V) Then Debug.Print "RgKeyVal: Dup Col: " & V
            O.Add V, iCno
        End If
    Next
Set RgKeyAdrDicH = O
End Function

Function RgLasCnoFmEdge&(A As Range)
Dim MaxCno&: MaxCno = WsMaxCno(A.Parent)
Dim Sq(): Sq = A(1, MaxCno - A.Column).Value
Dim Dr(): Dr = SqDr(Sq, 1)
RgLasCnoFmEdge = DrLasNonBlankIdx(Dr) + 1
End Function

Function RgLo(A As Range) As ListObject
Set RgLo = RgWs(A).ListObjects.Add(xlSrcRange, A)
End Function

Function RgNCol&(A As Range)
RgNCol = A.Columns.Count
End Function

Function RgNRow&(A As Range)
RgNRow = A.Rows.Count
End Function

Function RgR(A As Range, R) As Range
Set RgR = RgRCRC(A, R, 1, R, RgNCol(A))
End Function

Function RgR1&(A As Range)
RgR1 = A.Row
End Function

Function RgR2&(A As Range)
RgR2 = A.Row + A.Rows.Count - 1
End Function

Function RgRC(A As Range, R, C) As Range
Set RgRC = A.Cells(R, C)
End Function

Function RgRCRC(A As Range, R1, C1, R2, C2) As Range
Set RgRCRC = RgWs(A).Range(RgRC(A, R1, C1), RgRC(A, R2, C2))
End Function

Function RgRR(A As Range, R1, R2) As Range
Set RgRR = RgRCRC(A, R1, 1, R2, RgNCol(A))
End Function

Function RgToStr$(Rg As Range)
On Error GoTo R
RgToStr = Rg.Parent.Name & "!" & Rg.Address
Exit Function
R: RgToStr = ErStr("RgToStr")
End Function

Function RgWb(A As Range) As Workbook
Set RgWb = WsWb(A.Parent)
End Function

Function RgWs(A As Range) As Worksheet
Set RgWs = A.Parent
End Function
