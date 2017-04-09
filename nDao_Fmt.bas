Attribute VB_Name = "nDao_Fmt"
Option Compare Database
Option Explicit

Function FmtNmByRs$(NmStr$, Rs As DAO.Recordset)
FmtNmByRs = FmtNmByDic(NmStr$, RsDic(Rs))
End Function
