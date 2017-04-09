Attribute VB_Name = "nDao_DteTbl"
Option Compare Database
Option Explicit

Sub DteTblBld()
TblDrp "tblDte"
TblCrt "tblDte", "Dte Date, YY Byte, MM Byte, DD Byte, [Wk#] Byte, [Wk Day] Text(3)"
Dim mRs As DAO.Recordset:
Set mRs = CurrentDb.TableDefs("tblDte").OpenRecordset
With mRs
    Dim J%:
    For J = 0 To 10000
        .AddNew
        !Dte = #1/1/2017# + J
        !yy = Year(!Dte) - 2000
        !MM = Month(!Dte)
        !DD = Day(!Dte)
        .Fields("Wk#").Value = Format(!Dte, "ww")
        .Fields("Wk Day").Value = Format(!Dte, "ddd")
        .Update
    Next
    .Close
End With
End Sub
