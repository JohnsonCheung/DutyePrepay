Attribute VB_Name = "nXls_Fx"
Option Compare Database
Option Explicit

Function FxCpyAndOpn(oWb As Workbook, pFxFm$, pFxTo$, Optional OvrWrt As Boolean = False) As Boolean
Const cSub$ = "FxCpyAndOpn"
If VBA.Dir(pFxFm) = "" Then ss.A 1, "From file not exist": GoTo E

'If <OvrWrt>, delete <pFxTo> if exist, else prompt to overwrite if exist.
If VBA.Dir(pFxTo) <> "" Then
    If OvrWrt Then
        If Dlt_Fil(pFxTo) Then
            Dim mMsg$: mMsg = "Target Xls file [" & pFxTo & "] cannot be overwritten (or killed)||" & _
                "Check:|" & _
                "1. Check if the Target Xls is openned.  If is openned, close it and re-run||" & _
                "2. Otherwise, do following:|" & _
                "   1 Close all Xls files|" & _
                "   2 Press [Ctrl]+[Alt]+[Delete], Click [Task Manager] button|" & _
                "   3 A window [Windows Task Manager] is displayed.  Click [Processes] page ta|" & _
                "   4 Click the column [Image Name] to sort [Image Name]|" & _
                "   5 If there is any [Excel.exe] in the column [Image Name], highlight it [Excel.Exe] and Click [End Process]|" & vbLf & _
                "   6 Repeat [5] until no more [Excel.exe] in the column [Image Name]|" & _
                "   7 Re-run the program"
            ss.A 2, mMsg: GoTo E
        End If
    Else
        ss.A 3, "To file exist": GoTo E
    End If
End If
'Copy <pFxFm> to <pFxTo> and open <pFxTo> in mWb
If Cpy_Fil(pFxFm, pFxTo) Then ss.A 4: GoTo E
gXls.AutomationSecurity = msoAutomationSecurityForceDisable
Set oWb = gXls.Workbooks.Open(pFxTo, UpdateLinks:=False, IgnoreReadOnlyRecommended:=True)
Exit Function
R: ss.R
E: FxCpyAndOpn = True: ss.B cSub, cMod, "pFxFm,pFxTo,OvrWrt", pFxFm, pFxTo, OvrWrt
End Function

Sub FxCrtFmPthOfSngWsFx(Fx$, Pth$)
'Aim: Join all Xls files in {Pth} into one Xls {Fx}
'Assume: Each Xls in {Pth} has only 1 ws and have ws name and the file same being the same.
'==Start
Dim FfnTo$: FfnTo = Pth & Fx

'Create {mWbTo} by copy first Xls in {Pth} as {Fx}
Dim AyFnXls$(): AyFnXls = PthFnAy(Pth, "*.xls?")
If Sz(AyFnXls) = 0 Then Er "FxCrtFmPthOfSngWsFx: No xls in {Pth}", Pth
FileSystem.FileCopy Pth & AyFnXls(LBound(AyFnXls)) & ".xls", FfnTo
Dim mWbTo As Workbook: Set mWbTo = gXls.Workbooks.Open(FfnTo)

gXls.DisplayAlerts = False
'Loop each Xls file started from 2nd Xls in {Pth}
Dim iFnXls$, J As Byte
For J = 1 To UBound(AyFnXls)
    iFnXls = AyFnXls(J)
    Dim mWbFm As Workbook:
        If IsSingleWsXls(Pth & iFnXls, mWbFm) Then Stop

    ''Copy the {mFmWs} to {mWbTo}, then close mWbFm
    Dim mWs As Worksheet: If Crt_Ws_FmWs(mWs, mWbFm.Worksheets(1), , mWbTo) Then Stop
    mWbFm.Close
Next
Dlt_Dir Pth, "*.xls"
'Save mWbTo
gXls.DisplayAlerts = False
mWbTo.SaveAs Pth & Fx
gXls.DisplayAlerts = True
End Sub

Function FxCrtFmTpWithRfh(Fx$, FxTp$, Optional Vis As Boolean = False) As Workbook
FfnCpy FxTp, Fx, OvrWrt:=True
Dim O As Workbook
Set O = WbNew(Fx, Vis)
WbRfh O
Set FxCrtFmTpWithRfh = O
End Function

Function FxFstWsNm$(Fx)
Dim W As Workbook: Set W = FxWb(Fx)
FxFstWsNm$ = WbWs(W, 1).Name
WbCls W
End Function

Function FxIsSingleWs(Fx$) As Boolean
Dim W As Workbook
    Set W = FxWb(W)
FxIsSingleWs = W.Sheets.Count = 1
WbCls W, NoSav:=True
End Function

Function FxWb(Fx) As Workbook
FfnAsstExist Fx, "FxWb"
Dim X As Excel.Application: Set X = Appx
Dim W As Workbook
For Each W In X.Workbooks
    If W.FullName = Fx Then Set FxWb = W: Exit Function
Next
Set FxWb = X.Workbooks.Open(Fx, UpdateLinks:=False)
End Function

Function FxWrtPdf(Pfx$, Optional pFfnPDF$ = "", Optional pKeepXls As Boolean = False) As Boolean
Const cSub$ = "FxWrtPdf"
Dim mWb As Workbook: If Opn_Wb_R(mWb, Pfx) Then ss.A 1: GoTo E
Dim mFfnn$: mFfnn = Cut_Ext(Pfx)
Dim mFfnPDF$: mFfnPDF = Fct.NonBlank(pFfnPDF, mFfnn & ".pdf")
Dim mFfnPS$: mFfnPS = mFfnn & ".ps"
On Error GoTo R
If Set_PdfPrt(True) Then ss.A 2: GoTo E
mWb.PrintOut , , , , , True, , mFfnPS
If Set_PdfPrt(False) Then ss.A 3: GoTo E
Cls_Wb mWb, False
If Not pKeepXls Then Dlt_Fil Pfx
FxWrtPdf = Crt_PDF_FmFfnPS(mFfnPS, mFfnPDF)
Exit Function
R: ss.R
E: FxWrtPdf = True: ss.B cSub, cMod, "pFx,pFfnPdf", Pfx, pFfnPDF
End Function

Function FxWrtPdf__Tst()
'FxWrtPdf_Tst = FxWrtPdf("M:\07 ARCollection\ARCollection\PgmDoc.xls")
End Function

