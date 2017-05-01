Attribute VB_Name = "nDao_nTbl_nChk_Tbl"
Option Compare Database
Option Explicit

Sub TblAsstDupKey(T, Optional FstNFld% = 1, Optional A As database)
ErAsst TblChkDupKey(T, FstNFld%, A)
End Sub

Function TblAsstDupKey__Tst()
Dim T$: T = TmpNm

TblAsstDupKey T, 1
TblAsstDupKey T, 2
End Function

Sub TblAsstEmptyRec(T$, Optional A As database)
ErAsst TblChkEmptyRec(T, A)
End Sub

Sub TblAsstEmptyRec__Tst()
Const T$ = "#Tmp"
Const F$ = "AA BB"
TblDrp T
TblCrt T, "AA Text, BB Integer"
TblIns T, F, ApAv("1", 2)
TblIns T, F, ApAv("2", 3)
TblIns T, F, ApAv(Null, Null)
TblAsstEmptyRec T
TblDrp T
End Sub

Sub TblAsstFnStr(T$, FnStr$, Optional A As database)
ErAsst TblFldChkFnStr(T, FnStr, A)
End Sub

Sub TblAsstFnStr__Tst()
Const T$ = "#Tmp"
Const F$ = "AA BB"
TblDrp T
TblCrt T, "AA Text, BB Integer"
TblAsstFnStr T, "AA BB CC"
TblDrp T
End Sub

Function TblChkDupKey(T, Optional FstNFld% = 1, Optional A As database) As Variant()
'Aim: Chk if first {NPk} fields in {pNmt} has duplicate.  Return true is there is duplicate
Dim Sel$: Sel = Jn(TblFny(T, FstNFld, A), ",")
Dim Sql$: Sql = FmtQQ("Select Distinct ?, Count(*) as Cnt from [?] group by {0} having Count(*)>1", Sel, T)
TblChkDupKey = SqlDr(Sql, A)
End Function

Function TblChkEmptyRec(T$, Optional A As database) As Variant()
Dim D As database: Set D = DbNz(A)
Dim Rs As Recordset: Set Rs = TblRs(T, D)
Dim R&: R = 0
Dim RR&()
With Rs
    Dim Dr()
    While Not .EOF
        R = R + 1
        If RsIsEmptyRec(Rs) Then Push RR, R
        .MoveNext
    Wend
    .Close
End With
If AyIsEmpty(RR) Then Exit Function
Dim O$()
TblChkEmptyRec = ErNew("{Table} has {Rec#-List} being empty record, where Rec# starts counting from 1", T, Jn(RR, " "))
End Function
