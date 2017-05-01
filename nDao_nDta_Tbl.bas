Attribute VB_Name = "nDao_nDta_Tbl"
Option Compare Database
Option Explicit

Sub AA3()
TblDs__Tst
End Sub

Function TblDrAy(T, Optional FstNFld%, Optional A As database) As Variant()
TblDrAy = RsDrAy(DbNz(A).TableDefs(T).OpenRecordset, FstNFld)
End Function

Function TblDs(Optional TnPrm, Optional A As database) As Ds
Dim Tny$(): Tny = TnPrmToTny(TnPrm, A)
TblDs = TnyDs(Tny, A)
End Function

Sub TblDs__Tst()
Dim Tny$(): Tny = DbTny
Tny = AyExcl(Tny, "TblHasNoRec_IgnoreEr")
DsWb TblDs(Tny)
End Sub

Function TblDt(T, Optional FstNFld%, Optional A As database) As Dt
TblDt = RsDt(DbNz(A).TableDefs(T).OpenRecordset, FstNFld, T)
End Function

Sub TblInsDt(T, Dt As Dt, Optional A As database)
If DtIsNoRec(Dt) Then Exit Sub
Dim Rs As Recordset: Set Rs = TblRs(T, A)
Dim RIdx&(): RIdx() = RsIdx(Rs, Dt.Fny)
Dim Dr
For Each Dr In Dt.DrAy
    Rs.AddNew
    DrUpdRs Dr, Rs, RIdx
    Rs.Update
Next
Rs.Close
End Sub
