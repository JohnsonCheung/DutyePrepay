Attribute VB_Name = "nXls_XNm"
Option Compare Database
Option Explicit

Function XNmCpyToCell(pWbSrc As Workbook, pXlsNmSrc$, pWsTar As Worksheet) As Boolean
'Aim: Copy the range as pointed by {pXlsNmSrc} in {pWbSrc} to {pWsTar}
Const cSub$ = "XNmCpyToCell"
On Error GoTo R
Dim mXlsNm As Excel.Name: Set mXlsNm = pWbSrc.Names(pXlsNmSrc)
Dim mRge As Range: Set mRge = mXlsNm.RefersToRange
mRge.Copy pWsTar.Range("A1")
Exit Function
R: ss.R
E: XNmCpyToCell = True: ss.B cSub, cMod, "pWbSrc,pXlsNmSrc,pWbTar", ToStr_Wb(pWbSrc), pXlsNmSrc, ToStr_Ws(pWsTar)
End Function

Function XNmCpyToCell__Tst()
Const cFxTar$ = "C:\aa.xls"
Const cFx$ = "p:\Workingdir\Meta Db.xls"
Dim mWbTar As Workbook: If Crt_Wb(mWbTar, cFxTar, True) Then Stop
Dim mWsTar As Worksheet: Set mWsTar = mWbTar.Sheets(1)
Dim mWbSrc As Workbook: If Opn_Wb_R(mWbSrc, cFx) Then Stop
If XNmCpyToCell(mWbSrc, "DefTbl", mWsTar) Then Stop
mWbTar.Application.Visible = True
End Function

Function XNmCpyToFx(pWbSrc As Workbook, pXlsNmSrc$, pFxTar$, Optional pNmWsTar$ = "", Optional OvrWrt As Boolean = False) As Boolean
Const cSub$ = "XNmCpyToFx"
'Aim: Copy the range as defined in {pXlsNmSrc} in {pWbSrc} to of {pNmWsTar} in {pFxTar}.  If {pNmWsTar} is '', use {pXlsNmSrc}
On Error GoTo R
Dim mWbTar As Workbook: If Crt_Wb(mWbTar, pFxTar, OvrWrt) Then ss.A 1: GoTo E
If Dlt_AllWs_Except1(mWbTar) Then ss.A 2: GoTo E

If pNmWsTar = "" Then pNmWsTar = pXlsNmSrc
Dim mWsTar As Worksheet: Set mWsTar = mWbTar.Sheets(1)
mWsTar.Name = pNmWsTar
If XNmCpyToCell(pWbSrc, pXlsNmSrc, mWsTar) Then ss.A 3: GoTo E
If Cls_Wb(mWbTar, True) Then ss.A 1: GoTo E
Exit Function
R: ss.R
E: XNmCpyToFx = True: ss.B cSub, cMod, "pWbSrc,pXlsNmSrc,pFxTar", ToStr_Wb(pWbSrc), pXlsNmSrc, pFxTar
End Function

Sub XNmCpyToFx__Tst()
Const cFxTar$ = "C:\aa.xls"
Const cFx$ = "P:\WorkingDir\META Db.xls"
Dim mWb As Workbook: If Opn_Wb(mWb, cFx, True) Then Stop: GoTo E
If XNmCpyToFx(mWb, "DefTbl", cFxTar, "Tbl", True) Then Stop: GoTo E
If Cls_Wb(mWb) Then Stop: GoTo E
If Opn_Wb(mWb, cFxTar, , , True) Then Stop: GoTo E
E:
End Sub
