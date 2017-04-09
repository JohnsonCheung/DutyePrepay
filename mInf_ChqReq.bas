Attribute VB_Name = "mInf_ChqReq"
Option Compare Database
Option Explicit

Function ChqReqFdr$()
ChqReqFdr = FbCurPth & "Cheque Request\"
End Function

Function ChqReqFx(PermitId&)
Dim PermitNo$: PermitNo = TblPermitNo(PermitId)
Dim Fn$: Fn = "(" & Format(PermitId, "00000") & ") " & PermitNo & ".xls"
ChqReqFx = ChqReqFdr & Fn
End Function
