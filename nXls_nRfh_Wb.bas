Attribute VB_Name = "nXls_nRfh_Wb"
Option Compare Database
Option Explicit

Sub WbRfh(A As Workbook)
'Aim: Use CurrentDb as source to refresh given {pWorkbooks} data.
WbRfhPc A
Dim Ws As Worksheet
For Each Ws In A.Sheets
    WsRfh Ws
Next
WbRfhCht A
End Sub

Sub WbRfhCht(A As Workbook)
Dim Cht As Excel.Chart
WbRfhMsgShw A, "Charts"
For Each Cht In A.Charts
    ChtRfh Cht
Next
End Sub

Function WbRfhMsgPfx$(A As Workbook)
WbRfhMsgPfx = "Refreshing Wb[" & A.Name & "] "
End Function

Sub WbRfhMsgShw(A As Workbook, ObjTy$)
StsShw WbRfhMsgPfx(A) & ObjTy & " ...."
End Sub

Sub WbRfhPc(A As Workbook)
WbRfhMsgShw A, "PivotCaches"
Dim Pc As PivotCache
For Each Pc In A.PivotCaches
    PcRfh Pc
Next
End Sub
