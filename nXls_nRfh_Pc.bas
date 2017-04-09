Attribute VB_Name = "nXls_nRfh_Pc"
Option Compare Database
Option Explicit

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
