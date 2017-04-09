Attribute VB_Name = "nXls_Ay"
Option Compare Database
Option Explicit

Sub AyPutCellH(Ay, Cell As Range)
SqPutCell AySqH(Ay), Cell
End Sub

Sub AyPutCellH__Tst()
'1 Declare
Dim Ay
Dim Cell As Range

'2 Assign
Ay = Array(1, 2, 4, 5)
Set Cell = WsNew.Cells(2, 2)

'3 Calling
AyPutCellH Ay, Cell
Cell.Application.Visible = True
Stop
WsCls Cell.Parent, NoSav:=True
End Sub

Sub AyPutCellV(Ay, Cell As Range)
SqPutCell AySqV(Ay), Cell
End Sub
