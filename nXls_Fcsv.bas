Attribute VB_Name = "nXls_Fcsv"
Option Compare Database
Option Explicit

Sub FcsvPt(Fcsv, Cell As Range, RowFnStr$, Col$, Optional Pag$)

End Sub

Function FcsvWrtFx(pFfnCsv$, Optional Pfx$ = "", Optional OvrWrt As Boolean = False, Optional pKeepCsv = False) As Boolean
Const cSub$ = "Csv2Xls"
'Aim: Cv {pFfnCsv} to {pFx}
If VBA.Dir(pFfnCsv) = "" Then ss.A 1, "{pFfnCsv} not exist": GoTo E
If Pfx = "" Then Pfx = Repl_Ext(pFfnCsv, ".xls")
If Ovr_Wrt(Pfx, OvrWrt) Then ss.A 1: GoTo E
Dim mWb As Workbook: If Opn_Wb_R(mWb, pFfnCsv) Then ss.A 2: GoTo E
mWb.SaveAs Pfx, XlFileFormat.xlWorkbookNormal
mWb.Close
If Not pKeepCsv Then Dlt_Fil pFfnCsv
Exit Function
R: ss.R
E: FcsvWrtFx = True: ss.B cSub, cMod, "pFfnCsv,pFx,OvrWrt,pKeepCsv", pFfnCsv, Pfx, OvrWrt, pKeepCsv
End Function

Function FcsvWrtFx__Tst()
Const cFfnCsv$ = "c:\aa.csv"
Const cFx$ = "c:\aa.xls"
Close
Dim mFno As Byte: If Opn_Fil_ForOutput(mFno, cFfnCsv, True) Then Stop
Print #mFno, "aa,bb,cc,dd"
Print #mFno, "1,23,4,6"
Print #mFno, "2,3223,1234,1/2/2007"
Print #mFno, "221,323,423,621"
Close #mFno
If FcsvWrtFx(cFfnCsv, , True) Then Stop
Dim mWb As Workbook: If Opn_Wb(mWb, cFx) Then Stop
mWb.Application.Visible = True
End Function
