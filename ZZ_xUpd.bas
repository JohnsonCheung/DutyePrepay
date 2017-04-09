Attribute VB_Name = "ZZ_xUpd"
Option Compare Text
Option Explicit
Option Base 0

Sub TblUpdTbl(Tar$, Src$, NKeyFld%, Optional A As database)
'Aim: Upd {pNmtSrc} to {TarTn} for records exist in both tables
'     assuming first {pNKFld} are common primary in both tables
Dim O$
O = SqlStrOfUpd1(Tar, Src, NKeyFld, , A)
DbRunSql O, A
End Sub
