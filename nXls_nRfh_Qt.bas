Attribute VB_Name = "nXls_nRfh_Qt"
Option Compare Database
Option Explicit

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
