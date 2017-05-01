Attribute VB_Name = "nXls_nObj_nCell_nInf_Cell"
Option Compare Database
Option Explicit

Sub CellAddCmt(Cell As Range, Cmt$, W%, H%)
Cell.AddComment Cmt
With Cell.Comment.Shape
    .Width = W
    .Height = H
End With
End Sub

Function CellReSz(Cell As Range, Sq) As Range
Dim R, C
    R = UBound(Sq, 1)
    C = UBound(Sq, 2)
Set CellReSz = RgRCRC(Cell, 1, 1, R, C)
End Function

