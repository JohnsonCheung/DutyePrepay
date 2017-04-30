Attribute VB_Name = "nXls_nAct_nRfh"
Option Compare Database
Option Explicit

Function LoRfh(A As ListObject)
On Error Resume Next
QtRfh A.QueryTable
End Function


Sub PcRfh(A As PivotCache)
If A.SourceType <> xlDatabase Then Exit Sub
A.Connection = CnnStrFbOle(CurrentDb.Name)
A.BackgroundQuery = False
A.MissingItemsLimit = xlMissingItemsNone
'If pLExpr <> "" Then
'    If .CommandType <> xlCmdSql Then ss.A 1, "Given Command Type must be Sql": GoTo E
'    If InStr(.CommandText, "where") > 0 Then ss.A 2, "Given Sql should have have where": GoTo E
'    .CommandText = .CommandText & jj.Cv_Str(pLExpr, " where ")
'End If
On Error Resume Next
A.Refresh
End Sub

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

Sub QtRfh(A As QueryTable)
On Error Resume Next
'If pLExpr <> "" Then
'    If .CommandType <> xlCmdSql Then ss.A 4, "Given Command Type must be Sql": GoTo E
'    If InStr(.CommandText, "where") > 0 Then ss.A 5, "Given Sql should have have where": GoTo E
'    .CommandText = .CommandText & " WHERE " & pLExpr
'End If
A.Connection = CnnStrFbOle(CurrentDb.Name)
A.BackgroundQuery = False
A.Refresh False
End Sub


Sub ChtRfh(A As Chart)
On Error Resume Next
If IsNothing(A.PivotLayout) Then Exit Sub
PtRfh A.PivotLayout.PivotTable
End Sub


Sub ChtObjRfh(A As ChartObject)
On Error Resume Next
If IsNothing(A.Chart) Then Exit Sub
ChtRfh A.Chart
End Sub


Sub PtRfh(A As PivotTable)
On Error Resume Next
A.RefreshTable
End Sub


Sub WsRfh(A As Worksheet)
WsRfhLo A
WsRfhQt A
WsRfhPt A
WsRfhChtObj A
End Sub

Sub WsRfhChtObj(A As Worksheet)
Dim CObj As ChartObject
WsRfhMsgShw A, "Chart Objects"
For Each CObj In A.ChartObjects
    ChtObjRfh CObj
Next
End Sub

Sub WsRfhLo(A As Worksheet)
Dim LO As ListObject
WsRfhMsgShw A, "ListObjects"
For Each LO In A.ListObjects
    LoRfh LO
Next
StsClr
End Sub

Function WsRfhMsg$(A As Worksheet, ObjTy$)
WsRfhMsg = WbRfhMsgPfx(A.Parent) & "Ws[" & A.Name & "] " & ObjTy & " ...."
End Function

Sub WsRfhMsgShw(A As Worksheet, ObjTy$)
StsShw WsRfhMsg(A, ObjTy)
End Sub

Sub WsRfhPt(A As Worksheet)
Dim Pt As PivotTable
WsRfhMsgShw A, "PivotTables"
For Each Pt In A.PivotTables
    PtRfh Pt
Next
StsClr
End Sub

Sub WsRfhQt(A As Worksheet)
Dim Qt As QueryTable
WsRfhMsgShw A, "QueryTables"
For Each Qt In A.QueryTables
    QtRfh Qt
Next
StsClr
'If pLExpr <> "" Then
'    If .CommandType <> xlCmdSql Then ss.A 4, "Given Command Type must be Sql": GoTo E
'    If InStr(.CommandText, "where") > 0 Then ss.A 5, "Given Sql should have have where": GoTo E
'    .CommandText = .CommandText & " WHERE " & pLExpr
'End If
End Sub
