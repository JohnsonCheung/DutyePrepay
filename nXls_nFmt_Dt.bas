Attribute VB_Name = "nXls_nFmt_Dt"
Option Compare Database
Option Explicit

Sub DtPutCellWithFmt(Dt As Dt, Cell As Range, Fmtr As ListObjFmtr)
Dim OSq
Dim OListObj As ListObject
'Dim ODtaSq
'Dim ODtaRg As Range
'SqPutCell ODtaSq, ODtaRg
'    ODtaSq = DrAySq()
'
'    Set ODtaRg = RgRC(Cell, NHdrLin, 1)
'
'=======
SqPutCell OSq, Cell
ListObjFmt OListObj, Fmtr
End Sub
