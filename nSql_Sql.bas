Attribute VB_Name = "nSql_Sql"
Option Compare Database
Option Explicit

Sub SqlAsg(Sql$, ParamArray OAp())
Dim Dr(): Dr = SqlDr(Sql)
Dim Av(): Av = OAp
Dim N%: N = Sz(Av)
OAp(0) = Dr(0): If N = 1 Then Exit Sub
OAp(1) = Dr(1): If N = 2 Then Exit Sub
OAp(2) = Dr(2): If N = 3 Then Exit Sub
OAp(3) = Dr(3): If N = 4 Then Exit Sub
OAp(4) = Dr(4): If N = 5 Then Exit Sub
OAp(5) = Dr(5): If N = 6 Then Exit Sub
OAp(6) = Dr(6): If N = 7 Then Exit Sub
OAp(7) = Dr(7): If N = 8 Then Exit Sub
OAp(8) = Dr(8): If N = 9 Then Exit Sub
OAp(9) = Dr(9): If N = 10 Then Exit Sub
Stop
End Sub

Sub SqlBrw(Sql$, Optional QryNm$ = "qry")
'Aim: Create or set sql of given qQryNam and open to preview.  Usually for debug
QryCrt QryNm, Sql
DoCmd.OpenQuery QryNm, acViewPreview, acReadOnly
End Sub

Function SqlCol(Sql$, Optional A As database) As Variant()
Dim OAy()
SqlCol = SqlIntoAy(Sql, OAy, A)
End Function

Function SqlDr(Sql$, Optional A As database) As Variant()
SqlDr = FldsDr(SqlRs(Sql, A).Fields)
End Function

Function SqlDrAy(Sql$, Optional A As database) As Variant()
SqlDrAy = RsDrAy(SqlRs(Sql, A))
End Function

Function SqlDt(Sql$, Optional A As database) As Dt
SqlDt = RsDt(SqlRs(Sql, A))
End Function

Function SqlInt%(Sql$, Optional A As database)
With SqlRs(Sql, A)
    SqlInt = .Fields(0).Value
    .Close
End With
End Function

Function SqlIntoAy(Sql$, OAy, Optional A As database)
Erase OAy
With SqlRs(Sql, A)
    While Not .EOF
        If Not IsNull(.Fields(0).Value) Then Push OAy, .Fields(0).Value
        .MoveNext
    Wend
    .Close
End With
SqlIntoAy = OAy
End Function

Function SqlIntoAy__Tst()
Dim Sy$(): Sy = SqlIntoAy("Select PermitNo From Permit", Sy)
AyDmp Sy
End Function

Function SqlIsAny(Sql$, Optional A As database) As Boolean
SqlIsAny = Not SqlRs(Sql, A).EOF
End Function

Function SqlLng&(Sql$, Optional A As database)
SqlLng = SqlV(Sql, A)
End Function

Function SqlLngAy(Sql$, Optional A As database) As Long()
Dim O&()
SqlLngAy = SqlIntoAy(Sql, O, A)
End Function

Function SqlOptCur(Sql$, Optional A As database) As OptCur

End Function

Function SqlOptStr(Sql$, Optional A As database) As OptStr

End Function

Function SqlOptV(Sql$, Optional A As database) As OptV

End Function

Function SqlRs(Sql$, Optional A As database) As Recordset
Set SqlRs = DbNz(A).OpenRecordset(Sql)
End Function

Sub SqlRun(Sql$)
DoCmd.SetWarnings False
DoCmd.RunSql Sql
End Sub

Sub SqlRunAy(SqlAy$(), Optional A As database)
Dim I
For Each I In SqlAy
    DbNz(A).Execute I
Next
End Sub

Sub SqlRunQQ(SqlQQ$, ParamArray Ap())
Dim Av(): Av = Ap
Dim S$: S = FmtQQAv(SqlQQ, Av)
SqlRun S
End Sub

Function SqlSy(Sql$, Optional A As database) As String()
Dim O$()
SqlSy = SqlIntoAy(Sql, O, A)
End Function

Function SqlSy__Tst()
Dim Sy$(): Sy = SqlSy("Select PermitNo From Permit")
AyDmp Sy
End Function

Function SqlV(Sql$, Optional A As database)
With SqlRs(Sql, A)
    SqlV = .Fields(0).Value
    .Close
End With
End Function

Function Sqs$(Sql$, Optional A As database)
Sqs = SqlV(Sql, A)
End Function
