Attribute VB_Name = "nDao_Fmt"
Option Compare Database
Option Explicit

Function FmtRs$(NmMacro$, Rs As DAO.Recordset)
FmtRs = FmtDic(NmMacro, RsDic(Rs))
End Function
