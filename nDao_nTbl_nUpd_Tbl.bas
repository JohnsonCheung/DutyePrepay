Attribute VB_Name = "nDao_nTbl_nUpd_Tbl"
Option Compare Database
Option Explicit

Sub TblUpdFld(T$, KeyFld$, KeyVal$, FldToUpd$, V, Optional A As database)
Dim Sql$
Dim Rs As DAO.Recordset
    Dim Where$
    Where = KeyFld & "='" & KeyVal & "'"
    Sql = SqlStrOfSel(T, FldToUpd, Where)
    Set Rs = CurrentDb.OpenRecordset(Sql)
With Rs
    If .AbsolutePosition = -1 Then Er "No record in {Table} with given {KeyFld}={KeyVal}", T, KeyFld, KeyVal
    .Edit
    .Fields(0).Value = V
    .Update
    .Close
End With
End Sub

