Attribute VB_Name = "mInf_TblSku"
Option Compare Database
Option Explicit

Function TblSkuDesOpt(Sku) As OptStr
TblSkuDesOpt = SqlOptStr(FmtQQ("Select `Sku Description` from qSku where Sku = '?'", Sku))
End Function

Function TblSkuRatOpt(Sku) As OptCur
TblSkuRatOpt = SqlOptCur(FmtQQ("Select DutyRate from qSku where Sku='?'", Sku))
End Function
