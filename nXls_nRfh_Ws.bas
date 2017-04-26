Attribute VB_Name = "nXls_nRfh_Ws"
Option Compare Database
Option Explicit

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

Sub WsVis(A As Worksheet)
A.Application.Visible = True
End Sub
