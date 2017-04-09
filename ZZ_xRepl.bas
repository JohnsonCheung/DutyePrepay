Attribute VB_Name = "ZZ_xRepl"

'Option Compare Text
'Option Explicit
'Option Base 0
'Const cMod$ = cLib & ".xRepl"
'Function Repl_Ws_In2Wb(pWbTar As Workbook, pWbSrc As Workbook, pNmWs$) As Boolean
''Aim: replace the {pWs} in {pWbTar} by {WbSrc}
'Const cSub$ = "Repl_Ws_In2Wb"
'On Error GoTo R
'Dim mWsSrc As Worksheet: If Fnd_Ws(mWsSrc, pWbSrc, pNmWs) Then ss.A 1: GoTo E
'Dim mWsTar As Worksheet: If Fnd_Ws(mWsTar, pWbTar, pNmWs) Then If Add_Ws(mWsTar, pWbTar, pNmWs) Then ss.A 2: GoTo E
'If Repl_Ws(mWsTar, mWsSrc) Then ss.A 3: GoTo E
'Exit Function
'R: ss.R
'E: Repl_Ws_In2Wb = True: ss.B cSub, cMod, "pWbTar,pWbSrc,pNmWs", ToStr_Wb(pWbTar), ToStr_Wb(pWbSrc), pNmWs
'End Function

'Function Repl_Ws_In2Wb__Tst()
'Const cFx1$ = "c:\tmp\aa.xls"
'Const cFx2$ = "c:\tmp\bb.xls"
'Dim mWb1 As Workbook: If Crt_Wb(mWb1, cFx1, True) Then Stop: GoTo E
'Dim mWb2 As Workbook: If Crt_Wb(mWb2, cFx2, True) Then Stop: GoTo E
'If Cls_Wb(mWb2, True) Then Stop: GoTo E
'If Opn_Wb(mWb2, cFx2, True) Then Stop: GoTo E
'
'Dim mWs1 As Worksheet: Set mWs1 = mWb2.Sheets(1)
'Dim mWs2 As Worksheet: Set mWs2 = mWb2.Sheets(1)
'mWs2.Range("A1").Value = "From"
'If Repl_Ws_In2Wb(mWb1, mWb2, "ToBeDelete") Then Stop: GoTo E
'mWb1.Application.Visible = True
'Stop
'GoTo X
'E: Repl_Ws_In2Wb_Tst = True
'X: Cls_Wb mWb1
'   Cls_Wb mWb2
'End Function

'Function Repl_Ws_InFx(pFx$, pNmWsTar$, pNmWsSrc$) As Boolean
''Aim: replace the {pNmWsTar} by {pNmWsSrc} in same {pFx} and delete {pNmWsSrc}
'Const cSub$ = "Repl_Ws_InFx"
'Dim mWb As Workbook: If Opn_Wb_RW(mWb, pFx) Then ss.A 1: GoTo E
'If Repl_Ws_InWb(mWb, pNmWsTar, pNmWsSrc) Then ss.A 2: GoTo E
'If Cls_Wb(mWb, True) Then ss.A 3: GoTo E
'Exit Function
'R: ss.R
'E: Repl_Ws_InFx = True: ss.B cSub, cMod, "pFx,pNmWsTar,pNmWsSrc", pFx, pNmWsTar, pNmWsSrc
'End Function

'Function Repl_Ws_InFx__Tst()
'Dim mFx$: mFx = "c:\tmp\aa.xls"
'Dim mWs As Worksheet, mWb As Workbook
'If Crt_Wb(mWb, mFx, True, "Sheet1") Then Stop
'If Add_Ws_ByLnWs(mWb, "Sheet2,Sheet3,Sheet4") Then Stop
'Set mWs = mWb.Sheets("Sheet1"): If Set_Ws_ByLpAp(mWs, 1, "Sheet#,bbb,xxx", "Sheet1", 11111, 111119) Then Stop
'Set mWs = mWb.Sheets("Sheet2"): If Set_Ws_ByLpAp(mWs, 1, "Sheet#,bbb,xxx", "Sheet2", 22222, 222229) Then Stop
'Set mWs = mWb.Sheets("Sheet3"): If Set_Ws_ByLpAp(mWs, 1, "Sheet#,bbb,xxx", "Sheet3", 33333, 333339) Then Stop
'Set mWs = mWb.Sheets("Sheet4"): If Set_Ws_ByLpAp(mWs, 1, "Sheet#,bbb,xxx", "Sheet4", 44444, 444449) Then Stop
'If Cls_Wb(mWb, True) Then Stop
'MsgBox "Sheet2 will be replaced by Sheet4 and Sheet2 will be deleted", , "Repl_Ws"
'If Repl_Ws_InFx(mFx, "Sheet2", "Sheet4") Then Stop
'If Opn_Wb_R(mWb, mFx) Then Stop
'mWb.Application.Visible = True
'End Function

'Function Repl_Ws(pWsTar As Worksheet, pWsSrc As Worksheet) As Boolean
''Aim: replace the {pWsTar} by {pWsSrc} and delete {pWsTar}.  The 2 worksheets may in different wb.
''     If the workbook holding pWsSrc has only one worksheet, add a new ws will be added.
''     The pWsTar name will be preverse.
'Const cSub$ = "Repl_Ws"
'On Error GoTo R
'Dim mWb As Workbook: Set mWb = pWsSrc.Parent
'If mWb.Sheets.Count = 1 Then mWb.Sheets.Add
'Dim mNmWs$: mNmWs = pWsTar.Name
'pWsTar.Name = Format(Now, "yyyymmdd hhmmss")
'pWsSrc.Move After:=pWsTar
'If Dlt_Ws(pWsTar) Then ss.A 1: GoTo E
'Exit Function
'R: ss.R
'E: Repl_Ws = True: ss.B cSub, cMod, "pWsTar,pWsSrc", ToStr_Ws(pWsTar), ToStr_Ws(pWsSrc)
'End Function
'Function Repl_Ws_InWb(pWb As Workbook, pNmWsTar$, pNmWsSrc$) As Boolean
''Aim: replace the {pNmWsTar$} by {pNmWsTar} in {pWb} and delete {pNmWsTar}
'Const cSub$ = "Repl_Ws_InWb"
'On Error GoTo R
'Dim mWsTar As Worksheet: If Fnd_Ws(mWsTar, pWb, pNmWsTar) Then ss.A 1: GoTo E
'Dim mWsSrc As Worksheet: If Fnd_Ws(mWsSrc, pWb, pNmWsSrc) Then ss.A 2: GoTo E
'If Repl_Ws(mWsTar, mWsSrc) Then ss.A 3: GoTo E
'Exit Function
'R: ss.R
'E: Repl_Ws_InWb = True: ss.B cSub, cMod, "pWb,pNmWsTar,pNmWsTar", ToStr_Wb(pWb), pNmWsTar, pNmWsTar
'End Function

'Function Repl_Ws_InWb__Tst()
'Dim mFx$: mFx = "c:\tmp\aa.xls"
'Dim mWs As Worksheet, mWb As Workbook
'If Crt_Wb(mWb, mFx, True, "Sheet1") Then GoTo E
'If Add_Ws_ByLnWs(mWb, "Sheet2,Sheet3,Sheet4") Then GoTo E
'Set mWs = mWb.Sheets("Sheet1"): If Set_Ws_ByLpAp(mWs, 1, "Sheet#,bbb,xxx", "Sheet1", 11111, 111119) Then GoTo E
'Set mWs = mWb.Sheets("Sheet2"): If Set_Ws_ByLpAp(mWs, 1, "Sheet#,bbb,xxx", "Sheet2", 22222, 222229) Then GoTo E
'Set mWs = mWb.Sheets("Sheet3"): If Set_Ws_ByLpAp(mWs, 1, "Sheet#,bbb,xxx", "Sheet3", 33333, 333339) Then GoTo E
'Set mWs = mWb.Sheets("Sheet4"): If Set_Ws_ByLpAp(mWs, 1, "Sheet#,bbb,xxx", "Sheet4", 44444, 444449) Then GoTo E
'mWb.Application.Visible = True
'MsgBox "Sheet2 will be replaced by Sheet4 and Sheet2 will be deleted", , "Repl_Ws"
'If Repl_Ws_InWb(mWb, "Sheet2", "Sheet4") Then GoTo E
'Exit Function
'E: Repl_Ws_InWb_Tst = True
'End Function

'Function Repl_Cell_ByAy(Rg As Range, pFmVal$, pAyToVal$(), Optional pIsHDirection As Boolean = False) As Boolean
''Aim: replace the {pFmVal} in first cell of {Rg} by {pAyToVal} in either H or V direction
'Dim J%
'If pIsHDirection Then
'    For J = 0 To Siz_Ay(pAyToVal) - 1
'        Rg.Cells(1, 1 + J).Value = pAyToVal(J)
'    Next
'Else
'    For J = 0 To Siz_Ay(pAyToVal) - 1
'        Rg.Cells(1 + J, 1).Value = pAyToVal(J)
'    Next
'End If
'End Function
'Function Repl_WsChtTit(oChtTit As ChartTitle, pAyK$(), pAyV$()) As Boolean
'With oChtTit
'    If IsMacro(.Text) Then .Text = Fmt_Str_ByAyKV(.Text, pAyK, pAyV)
'End With
'End Function
'Function Repl_Ext$(pFfn$, pExt$)
'Dim p%: p = InStrRev(pFfn, ".")
'If p <= 0 Then Repl_Ext = pFfn$ & pExt
'Repl_Ext = Left(pFfn, p - 1) & pExt
'End Function
'Function Repl_WsPagSetup(oPagSetup As PageSetup, pAyK$(), pAyV$()) As Boolean
'With oPagSetup
'    If IsMacro(.LeftHeader) Then .LeftHeader = Fmt_Str_ByAyKV(.LeftHeader, pAyK, pAyV)
'    If IsMacro(.RightHeader) Then .RightHeader = Fmt_Str_ByAyKV(.RightHeader, pAyK, pAyV)
'    If IsMacro(.CenterHeader) Then .CenterHeader = Fmt_Str_ByAyKV(.CenterHeader, pAyK, pAyV)
'End With
'End Function
'Function Repl_Pfx_InAy(oAyTar$(), pPfxTar$, pAySrc$(), pPfxSrc$) As Boolean
'Const cSub$ = "Repl_Pfx_InAy"
'If pPfxTar = "" Then oAyTar = pAySrc: Exit Function
'Dim N%: N% = Siz_Ay(pAySrc): If N = 0 Then oAyTar = pAySrc: Exit Function
'ReDim oAyTar(N - 1)
'Dim L%: L = Len(pPfxSrc)
'Dim J%: For J = 0 To N - 1
'    If Left(pAySrc(J), L) <> pPfxSrc Then ss.A 1, "One of element in pAySrc does not have the pPfxSrc": GoTo E
'    oAyTar(J) = pPfxTar & mID(pAySrc(J), L + 1)
'Next
'Exit Function
'E: Repl_Pfx_InAy = True: ss.B cSub, cMod, "pPfxTar,pAySrc,pPfxSrc", pPfxTar, ToStr_Ays(pAySrc), pPfxSrc
'End Function
'Function Repl_Rge_ByAy(Rg As Range, pFmVal$, pAyToVal$(), Optional pIsHDirection As Boolean = False) As Boolean
'Const cSub$ = "Repl_Rge_Cell_ByAy"
''Aim: replace the {pFmVal} in {Rg} by {pAyToVal} in either H or V direction
'Dim mWs As Worksheet: Set mWs = Rg.Parent
'mWs.Outline.ShowLevels 8, 8
'
'Dim mRge As Range
'Set mRge = Rg.Find(What:=pFmVal, After:=Rg.Cells(1, 1), LookIn:=xlValues _
'    , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext _
'    , MatchCase:=False, SearchFormat:=False)
'
'While TypeName(mRge) <> "Nothing"
'    Repl_Cell_ByAy mRge, pFmVal, pAyToVal, pIsHDirection
'    Set mRge = mRge.FindNext
'Wend
'End Function
'Function Repl_Rge_ByAy__Tst()
'Const cFfnFm$ = "R:\Sales Simulation\Simulation\Templates\Topaz Data Import file ({StreamCode}).xls"
'Const cFfnTo$ = "c:\temp\a.xls"
'Dim mWb As Workbook: If FxCpyAndOpn(mWb, cFfnFm, cFfnTo) Then Stop
'Dim mWs As Worksheet: Set mWs = mWb.Sheets("SumTotalEuro {BrandGroupName}")
'Dim mAy$(9), J%
'For J = 0 To 9
'    mAy(J) = "Johnson-" & J
'Next
'If Repl_Rge_ByAy(mWs, "{BrandNameListDown}", mAy) Then Stop
'mWb.Application.Visible = True
'mWs.Activate
'End Function
'Function Repl_RgeVal(Rg As Range, pFmVal$, pToVal$) As Boolean
''Aim: Repl value in {Rg} from {pFmVal} to {pToVal}
'Const cSub$ = "Repl_RgeVal"
'On Error GoTo R
'Rg.Application.DisplayAlerts = False
'
'Dim mWs As Worksheet: Set mWs = Rg.Parent
'mWs.Outline.ShowLevels 8, 8
'
'Dim mCell As Range
'Set mCell = Rg.Find(What:=pFmVal, LookIn:=xlValues _
'    , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext _
'    , MatchCase:=False, SearchFormat:=False)
'While TypeName(mCell) <> "Nothing"
'    mCell.Value = Replace(mCell.Value, pFmVal, pToVal)
'    Set mCell = Rg.FindNext(mCell)
'Wend
'Rg.Application.DisplayAlerts = True
'Exit Function
'R: ss.R
'E: Repl_RgeVal = True: ss.B cSub, cMod
'End Function
'Function Repl_RgeVal__Tst()
'Const cFfnFm$ = "R:\Sales Simulation\Simulation\Templates\Topaz Data Import file ({StreamCode}).xls"
'Const cFfnTo$ = "c:\temp\a.xls"
'Dim mWb As Workbook: If FxCpyAndOpn(mWb, cFfnFm, cFfnTo) Then Stop
'Dim mWs As Worksheet: Set mWs = mWb.Sheets("SumTotalEuro {BrandGroupName}")
'If Repl_RgeVal(mWs.Cells, "{BrandGroupName}", "Johnson") Then Stop
'mWb.Application.Visible = True
'mWs.Activate
'End Function
'Function Repl_Sql(pPfx$, pFm$, pTo$) As Boolean
'Dim L%: L = Len(pPfx)
'Dim iQry As QueryDef: For Each iQry In CurrentDb.QueryDefs
'    If Left(iQry.Name, L) = pPfx Then
'        If InStr(iQry.Sql, pFm) > 0 Then
'            Debug.Print "replacing Qry ... "; iQry.Name
'            iQry.Sql = Replace(iQry.Sql, pFm, pTo)
'        End If
'    End If
'Next
'End Function
'Function Repl_Xls(pFx$, pNmtHdr$, Optional pNmtDet$ = "", Optional pNmDet$ = "") As Boolean
''Aim: Substitue the [variables] in {pFfnDoc}.  The variables are in format of {xxx} where xxx is the fields of the {pRsHdr} or {pRsDet}.
''     {pRsDet} are always fill in "Word's Table" having substring {<<pNmDet>>} in cell(1,1).  Each record in will be filled starting from 3rd row of the table.
''     The row of the "Word's Table" will be created automatically
'Const cSub$ = "Repl_Xls"
'Dim mRs As DAO.Recordset
'End Function
'Function Repl_Wrd(pFfnDoc$, pRsHdr As DAO.Recordset, Optional pRsDet As DAO.Recordset = Nothing, Optional pNmDet$ = "", Optional pFfnDetTp$, Optional pNHdrRows As Byte = 2) As Boolean
''Aim: Substitue the [variables] in {pFfnDoc}.  The variables are in format of {xxx} where xxx is the fields of the {pRsHdr} or {pRsDet}.
''     {pRsDet} are always fill in "Word's Table" having substring {<<pNmDet>>} in cell(1,1).  Each record in will be filled starting from 3rd row of the table.
''     The row of the "Word's Table" will be created automatically
'Const cSub$ = "Repl_Wrd"
'Dim mWrd As Word.Document: If Opn_Wrd_RW(mWrd, pFfnDoc) Then ss.A 1: GoTo E
'Dim iFld As DAO.Field
'Dim mFnd As Word.Find: Set mFnd = mWrd.Range.Find
'
''With mFnd
''    .Forward = False
''    .ClearFormatting
''    .MatchWholeWord = False
''    .MatchCase = False
''    .Wrap = wdFindContinue
''End With
'gWrd.ActiveWindow.ActivePane.View.Type = wdPrintView
''
''gWrd.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekPrimaryHeader
''For Each iFld In pRsHdr.Fields
''    mFnd.Execute "{" & iFld.Name & "}", False, False, , , , False, , , Nz(iFld.Value, ""), WdReplace.wdReplaceAll
''Next
''gWrd.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekPrimaryFooter
''For Each iFld In pRsHdr.Fields
''    mFnd.Execute "{" & iFld.Name & "}", False, False, , , , False, , , Nz(iFld.Value, ""), WdReplace.wdReplaceAll
''Next
''gWrd.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageFooter
''For Each iFld In pRsHdr.Fields
''    mFnd.Execute "{" & iFld.Name & "}", False, False, , , , False, , , Nz(iFld.Value, ""), WdReplace.wdReplaceAll
''Next
''gWrd.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageHeader
''For Each iFld In pRsHdr.Fields
''    mFnd.Execute "{" & iFld.Name & "}", False, False, , , , False, , , Nz(iFld.Value, ""), WdReplace.wdReplaceAll
''Next
'gWrd.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekMainDocument
'For Each iFld In pRsHdr.Fields
'    mFnd.Execute "{" & iFld.Name & "}", False, False, , , , False, , , Nz(iFld.Value, ""), WdReplace.wdReplaceAll
'Next
'
''-- Find if Detail Table exist ---------
'If pNmDet = "" Then GoTo NoDet
'If IsNothing(pRsDet) Then GoTo NoDet
'
'Dim mCase As Byte
'mCase = 2
'Select Case mCase
'Case 1
'    Dim iTbl As Word.Table, mFound As Boolean
'    For Each iTbl In mWrd.Tables
'        If iTbl.Rows.Count <> 3 Then GoTo NxtTbl
'        If iTbl.Rows(1).Cells.Count <= 0 Then GoTo NxtTbl
'        If InStr(iTbl.Rows(1).Cells(1).Range.Text, "{" & pNmDet & "}") = 0 Then GoTo NxtTbl
'        mFound = True: Exit For
'NxtTbl:
'    Next
'    If Not mFound Then GoTo NoDet
'    '-- Replace {<<pNmDet>>} to empty
'    mFnd.Execute "{" & pNmDet & "}", False, False, , , , False, , , "", WdReplace.wdReplaceAll
'    '-- Detail ---------
'    With pRsDet
'        iTbl.Rows(3).Select
'        mWrd.Application.Selection.Copy
'        While Not .EOF
'            mWrd.Application.Selection.Paste
'            .MoveNext
'        Wend
'        iTbl.Rows(3).Delete
'        .MoveFirst
'
'        Dim iRec%: iRec = 0
'        While Not .EOF
'            For Each iFld In pRsDet.Fields
'                With iTbl.Rows(3 + iRec).Range.Find
'                    .Forward = False
'                    .ClearFormatting
'                    .MatchWholeWord = False
'                    .MatchCase = False
'                    .Wrap = wdFindStop
'                    .Execute "{" & iFld.Name & "}", , , , , , , , , Nz(iFld.Value, ""), WdReplace.wdReplaceOne
'                End With
'            Next
'            iRec = iRec + 1
'            .MoveNext
'        Wend
'    End With
'Case 2
'    If VBA.Dir(pFfnDetTp) = "" Then ss.A 3, "Template file for Detail Records does not exist": GoTo E
'    Dim mWb As Workbook ' The Tp WB needs to keep open so that the format can be copied from source clip board
'    '
'    Stop
'    'If Crt_Clip_ByRs(pFfnDetTp$, 3, pRsDet, mWb) Then ss.A 2:Goto E
'    With mWrd.Application.Selection.Find
'        .ClearFormatting
'        .Text = "{" & pNmDet & "}"
'        .Replacement.Text = ""
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .MatchCase = False
'        .MatchWholeWord = False
'        .MatchByte = False
'        .MatchWildcards = False
'        .MatchSoundsLike = False
'        .MatchAllWordForms = False
'        If .Execute Then mWrd.Application.Selection.Paste
'        Cls_Wb mWb
'    End With
'    'Assume there is only one table
'    Dim iRow%
'    For iRow = 1 To pNHdrRows
'        mWrd.Tables(1).Rows(iRow).HeadingFormat = True
'    Next
'End Select
'
'NoDet:
'    If DocCls(mWrd, True) Then ss.A 3: GoTo E
'Exit Function
'R: ss.R
'E: Repl_Wrd = True: ss.B cSub, cMod, "pFfnDoc,pRsHdr,pRsDet,pNmDet,pFfnDetTp,pNHdrRows", pFfnDoc, ToStr_Rs_NmFld(pRsHdr), ToStr_Rs_NmFld(pRsDet), pNmDet, pFfnDetTp, pNHdrRows
'End Function

'Function Repl_Wrd__Tst()
'Const cFfn$ = "c:\aa.doc"
''Dim mFfnTp$: mFfnTp = "C:\DOC1.DOC"
'Dim mFbOldQsTmp$: If Fnd_Sffn_LgcMdbTmp(mFbOldQsTmp, "GenRmd") Then Stop
'If TblCrt_FmLnkLnt(mFbOldQsTmp, "tmpBldOneRmd_Hdr,tmpBldOneRmd_Det") Then Stop
'Dim mFfnTp$: mFfnTp = "M:\07 ARCollection\ARCollection\WorkingDir\Templates\Template_ReminderLvl3(English).doc"
'Dim mRsHdr As DAO.Recordset: Set mRsHdr = CurrentDb.TableDefs("tmpBldOneRmd_Hdr").OpenRecordset
'Dim mRsDet As DAO.Recordset: Set mRsDet = CurrentDb.TableDefs("tmpBldOneRmd_Det").OpenRecordset
'If Cpy_Fil(mFfnTp, cFfn) Then Stop
'If Repl_Wrd(cFfn, mRsHdr, mRsDet, "InvDet", Sffn_Tp("RmdInvDet(English)")) Then Stop
'gWrd.Documents.Open cFfn
'gWrd.Visible = True
'End Function

'Function Repl_WsChtObj(oWs As Worksheet, pAyK$(), pAyV$()) As Boolean
'Const cSub$ = "Repl_WsChtObj"
'Dim iChtObj As ChartObject
'For Each iChtObj In oWs.ChartObjects
'    If Repl_WsChtTit(iChtObj.Chart.ChartTitle, pAyK, pAyV) Then ss.A 2: GoTo E
'Next
'Exit Function
'R: ss.R
'E: Repl_WsChtObj = True: ss.B cSub, cMod, "pAyK,pAyV", ToStr_Ays(pAyK), ToStr_Ays(pAyV)
'End Function
'

