Attribute VB_Name = "nXls_nObj_nWb_nDo_Wb"
Option Compare Database
Option Explicit

Sub WbDltWs(A As Workbook, WsIdx)
Dim S As Boolean: S = A.Application.DisplayAlerts
A.Application.DisplayAlerts = False
WbWs(A, WsIdx).Delete
A.Application.DisplayAlerts = S
End Sub

Sub WbDltWsExcp(A As Workbook, ExcpWsIdx)
If Not WbHasWs(A, ExcpWsIdx) Then Er "WbDltWsExcp: {Wb} does not have {ExcpWsIdx}", WbToStr(A), ExcpWsIdx
Dim Nm$: Nm = WbWs(A, ExcpWsIdx).Name
While A.Sheets.Count >= 2
    If A.Sheets(1).Name = Nm Then
        WbDltWs A, 2
    Else
        WbDltWs A, 1
    End If
Wend
End Sub

Function WbDltWsExcpt__Tst()
Dim Wb As Workbook: Set Wb = WbNew
Wb.Sheets.Add
Wb.Sheets.Add
Wb.Sheets.Add
Wb.Application.Visible = True
WbDltWsExcp Wb, "Sheet2"
Stop
WbClsNosav Wb
End Function

Sub WbDltWsIfNeed(A As Workbook, WsIdx)
If WbHasWs(A, WsIdx) Then WbDltWs A, WsIdx
End Sub

Sub WbNewXNm(A As Workbook, XNm$, ReferToAdr$)
A.Names.Add XNm, ReferToAdr$
End Sub

Sub WbSetPrp(A As Workbook, Optional Tit$, Optional Subj$, Optional Author$, Optional Comments$, Optional Keywords$)
WbSetPrp_One A, "Title", Tit
WbSetPrp_One A, "Subject", Subj
WbSetPrp_One A, "Author", Author
WbSetPrp_One A, "Comments", Comments
WbSetPrp_One A, "Keywords", Keywords
End Sub

Sub WbSetPrp_One(A As Workbook, PrpNm$, V)
A.BuiltinDocumentProperties(PrpNm).Value = V
End Sub
