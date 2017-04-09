Attribute VB_Name = "ZZ_xFnd"

'Option Compare Text
'Option Explicit
'Option Base 0
'Const cMod$ = cLib & ".xFnd"
'Function Fnd_MsgBoxSty(pTypMsg As eTypMsg) As VbMsgBoxStyle
'Dim mA As VbMsgBoxStyle
'Select Case pTypMsg
'    Case eTypMsg.eCritical, eTypMsg.ePrmErr: mA = vbCritical
'    Case eTypMsg.eWarning: mA = vbExclamation
'    Case eTypMsg.eTrc, eTypMsg.eUsrInfo: mA = vbInformation
'    Case Else: mA = vbInformation
'End Select
'mA = mA Or vbDefaultButton1
'If SysCfg_IsDbg Then Fnd_MsgBoxSty = mA Or vbYesNo
'End Function

'
'Function Fnd_NxtFfn$(pFfn$)
''Aim: If pFfn exist, find next Ffn by adding (n) to the end of the file name.
'Const cSub$ = "Ffn_NxtFfn"
'If VBA.Dir(pFfn) = "" Then Fnd_NxtFfn = pFfn: Exit Function
'Dim mP%: mP = InStrRev(pFfn, ".")
'Dim mA$, mB$
'If mP = 0 Then
'    mA = pFfn
'Else
'    mA = Left(pFfn, mP - 1)
'    mB = mID(pFfn, mP)
'End If
'Dim J%
'For J = 0 To 100
'    Dim mFfn$: mFfn = mA & "(" & J & ")" & mB
'    If VBA.Dir(mFfn) = "" Then Fnd_NxtFfn = mFfn: Exit Function
'Next
'ss.A 1, "Quit impossible to reach here.. Having 100 next file exist"
'E: ss.B cSub, cMod, "pFfn", pFfn
'End Function
'Function Fnd_AyCnoImpFld(oAyCno() As Byte, oAmFld() As tMap, Rg As Range _
'    , Optional pRithNmFld% = -1, Optional pRithImp% = -4) As Boolean
''Aim: Whenever @ {pRithImp} & {pRithNmFld}, there is value a TypFld (vbString) & NmFld (vbString),
''     The column is a Import Field.  Put its name, type & Cno into {oAnFld}, {oAyTypFld} & {oAyCno}
'Const cSub$ = "Fnd_AyCnoImpFld"
'On Error GoTo R
'Clr_Am oAmFld
'Clr_AyByt oAyCno
'Dim mRnoImp&: mRnoImp = Rg.Row + pRithImp
'Dim mRnoNmFld&: mRnoNmFld = Rg.Row + pRithNmFld
'Dim mCnoBeg As Byte: mCnoBeg = Rg.Column
'
'Dim mWs As Worksheet: Set mWs = Rg.Parent
'Dim mV: mV = mWs.Cells(mRnoImp, mCnoBeg).Value
'If VarType(mV) <> vbString Then ss.A 1, Fmt_Str("Cell({0},{1}) must be vbString", mRnoImp, mCnoBeg), , "The Cell", mV: GoTo E
'If mV <> "Import:" & mWs.Name Then ss.A 1, Fmt_Str("Cell({0},{1}) must be 'Import:{2}'", mRnoImp, mCnoBeg, mWs.Name), , "The Cell", mV: GoTo E
'Dim iCno As Byte, N As Byte, mCnoLas As Byte: If Fnd_CnoLas(mCnoLas, Rg(0, 1)) Then ss.A 1: GoTo E
'If mCnoLas - Rg.Column < 1 Then ss.A 2, "There should at least 2 columns Id & Nam", , "Rg.Column,mCnoLas", Rg.Column, mCnoLas: GoTo E
'For iCno = mCnoBeg To mCnoLas
'    mV = mWs.Cells(mRnoNmFld, iCno).Value
'    Dim mT: mT = mWs.Cells(mRnoImp, iCno).Value
'    If VarType(mV) = vbString Then
'        Select Case VarType(mT)
'        Case vbString: ReDim Preserve oAmFld(N), oAyCno(N)
'                       oAmFld(N).F1 = mV: oAyCno(N) = iCno: oAmFld(N).F2 = mT
'                       N = N + 1
'        Case vbEmpty
'        Case Else:     ss.A 2, "It has a vbString field name, but a invalid Type", , "iCno,NmFld,TypFld", iCno, mV, mT
'        End Select
'    Else
'        If VarType(mT) <> vbEmpty Then ss.A 3, "A non field name cannot have non-empty Type", , "iCno,non-empty-type", iCno, mT
'    End If
'Next
'If InStr(oAmFld(0).F1, "_") > 0 Then
'    oAmFld(0).F2 = "Text 255"
'Else
'    oAmFld(0).F2 = "Long"
'End If
'Exit Function
'R: ss.R
'E: Fnd_AyCnoImpFld = True: ss.B cSub, cMod, "Rg,pRithImp", ToStr_Rge(Rg), pRithImp
'End Function

'Function Fnd_AyCnoImpFld__Tst()
'Dim mWb As Workbook: If Crt_Wb(mWb, "c:\tmp\bb.xls", True, "Sheet1") Then Stop: GoTo E
'
''^^
'mWb.Sheets(1).Cells(1, 1).Value = "Import:Sheet1"
'
'mWb.Sheets(1).Cells(4, 1).Value = "Id_Id"
'
'mWb.Sheets(1).Cells(1, 2).Value = "Text 1"
'mWb.Sheets(1).Cells(4, 2).Value = "Nm1"
'
'mWb.Sheets(1).Cells(1, 5).Value = "Text 2"
'mWb.Sheets(1).Cells(4, 5).Value = "Nm2"
'
'mWb.Sheets(1).Cells(1, 6).Value = "Text 3"
'mWb.Sheets(1).Cells(4, 6).Value = "Nm3"
'
'mWb.Sheets(1).Cells(1, 7).Value = "Text 4"
'mWb.Sheets(1).Cells(4, 7).Value = "Nm4"
'
'
'Dim mAmFld() As tMap, mAyCno() As Byte: If Fnd_AyCnoImpFld(mAyCno, mAmFld, mWb.Sheets(1).Range("A5")) Then Stop
'Debug.Print ToStr_LpAp(CtComma, "mAyCno", ToStr_AyByt(mAyCno))
'Debug.Print ToStr_LpAp(CtComma, "mAmFld", ToStr_Am(mAmFld, " "))
'Shw_DbgWin
'mWb.Application.Visible = True
'Stop
'Cls_Wb mWb, False, True
'E: Fnd_AyCnoImpFld_Tst = True
'X: Cls_Wb mWb, , True
'End Function

'Function Fnd_AyRecCnt(oAyRecCnt&(), pAnt$(), Optional pDb As database) As Boolean
''Aim: Find {oAyRecCnt} of each table of {pDb!pLnt}
'Const cSub$ = "Fnd_AyRecCnt"
'On Error GoTo R
'Dim N%, J%
'For J = 0 To Siz_Ay(pAnt) - 1
'    ReDim Preserve oAyRecCnt(N)
'    oAyRecCnt(N) = Fct.RecCnt(pAnt(J), pDb)
'    N = N + 1
'Next
'If N = 0 Then Clr_AyLng oAyRecCnt
'Exit Function
'R: ss.R
'E: Fnd_AyRecCnt = True: ss.B cSub, cMod, "pAnt,pDb", ToStr_Ays(pAnt), ToStr_Db(pDb)
'End Function
'Function Fnd_AyCnoColr(oAyCno() As Byte, oAyColr&(), Rg As Range, pRnoColrIdx&) As Boolean
''Aim: Find the color of a row {pRnoColrIdx} into {oAyCno} & {oAyColr}.  Row Rg(0,1) will be used to detect the start and end column
'Const cSub$ = "Fnd_AyCnoColr"
'On Error GoTo R
'Clr_AyLng oAyColr
'Clr_AyByt oAyCno
'Dim iCno As Byte, N%, mCnoLas As Byte: If Fnd_CnoLas(mCnoLas, Rg(0, 1)) Then ss.A 1: GoTo E
'Dim mWs As Worksheet: Set mWs = Rg.Parent
'For iCno = Rg.Column To mCnoLas
'    Dim mRge As Range: Set mRge = mWs.Cells(pRnoColrIdx, iCno)
'    Dim mColr&: mColr = mRge.Interior.Color
'    If mColr <> G.CtColrNone Then
'        ReDim Preserve oAyCno(N), oAyColr(N)
'        oAyCno(N) = iCno
'        oAyColr(N) = mRge.Interior.Color
'        N = N + 1
'    End If
'Next
'Exit Function
'R: ss.R
'E: Fnd_AyCnoColr = True: ss.B cSub, cMod, "Rg,pRnoColrIdx", ToStr_Rge(Rg), pRnoColrIdx
'End Function

'Function Fnd_AyCnoColr__Tst()
'Dim mWb As Workbook: If Crt_Wb(mWb, "c:\tmp\aa.xls", True, "Sheet1") Then Stop: GoTo E
'Dim mAyCno() As Byte, mAyColr&()
'mWb.Sheets(1).Cells(2, 5).Interior.Color = 123
'If Fnd_AyCnoColr(mAyCno, mAyColr, mWb.Sheets(1), 2) Then Stop
'Debug.Print ToStr_AyByt(mAyCno)
'Debug.Print ToStr_AyLng(mAyColr)
'Shw_DbgWin
'Stop
'Cls_Wb mWb, False, True
'E: Fnd_AyCnoColr_Tst = True
'End Function

'Function Fnd_AyRno_Visible(oAyRno&(), Rg As Range) As Boolean
''Aim: Find oAyRno() which is not hidden row started at {Rg}
'Const cSub$ = "Fnd_AyRno_Visible"
'On Error GoTo R
'Clr_AyLng oAyRno
'Dim iRno&, N%, mWs As Worksheet: Set mWs = Rg.Parent
'For iRno = Rg.Row To Rg(1, 1).End(xlDown).Row
'    If Not mWs.Rows(iRno).Hidden Then
'        ReDim Preserve oAyRno(N): oAyRno(N) = iRno: N = N + 1
'    End If
'Next
'Exit Function
'R: ss.R
'E: Fnd_AyRno_Visible = True: ss.B cSub, cMod, "Rg,pCno", ToStr_Rge(Rg)
'End Function
''Function Fnd_DocSml_ByRno(oDocSml As DOMDocument60, pWs As Worksheet, pRno&) As Boolean
'''Aim: find {oLnv}, which is a string of one line one Name=Value from {pRno} of {pWs} by using all Ws Names begins with x in {pWs}
''Clr_Doc oDocSml
''Dim N%, J%, mNod As MSXML2.IXMLDOMNode, mChd As IXMLDOMNode
''Set mChd = oDocSml.createNode(NODE_ELEMENT, "SML", ""): Set mNod = oDocSml.appendChild(mChd)
''Set mChd = oDocSml.createNode(NODE_ELEMENT, "Rec", ""): Set mNod = oDocSml.appendChild(mChd)
''For J = 0 To pWs.Names.Count - 1
''    With pWs.Names(J)
''        If Left(.Name, Len(pWs.Name) + 2) = pWs.Name & "!x" Then
''            Dim mNm$: mNm = mID(.Name, Len(pWs.Name) + 3)
''            Set mChd = oDocSml.createNode(NODE_ELEMENT, mNm, "")
''            mChd.Text = Nz(pWs.Cells(pRno, .RefersToRange.Column).Value, "")
''            mNod.appendChild mChd
''        End If
''    End With
''Next
''End Function
''Function Fnd_DocSml_ByFfn(oDocSml As MSXML2.DOMDocument60, pFfn$) As Boolean
'''Aim: Read content from {pFfn} and create {oDocSml}
''Const cSub$ = "Fnd_DocSml_ByFfn"
''Clr_Doc oDocSml
''oDocSml.Load pFfn
''If Chk_DocSml(oDocSml) Then ss.A 1: GoTo E
''Exit Function
''E: Fnd_DocSml_ByFfn = True: ss.B cSub, cMod, ToStr_Doc(oDocSml)
''End Function
'Function Fnd_Anq_wSubStr(oAnq$(), pSubStr$, Optional pDb As database, Optional pSilent) As Boolean
'Const cSub$ = "Fnd_Anq_wSubStr"
'On Error GoTo R
'Dim iQry As QueryDef
'For Each iQry In CurrentDb.QueryDefs
'    If Left(iQry.Name, 1) <> "~" Then
'        If InStr(iQry.Sql, pSubStr) > 0 Then Add_AyEle oAnq, iQry.Name
'    End If
'Next
'Exit Function
'R: ss.R
'E: Fnd_Anq_wSubStr = True: If Not pSilent Then ss.B cSub, cMod, "pStr"
'End Function
'Function Fnd_AnFld_ReqTxt(oAnFld$(), pNmt$, Optional pDb As database) As Boolean
''Aim: Find {oAnFld} of {pNmt} which is either (Text or Memo) is IsReq
'Const cSub$ = "Fnd_AnFld_ReqTxt"
'Dim mDb As database: Set mDb = DbNz(pDb)
'Dim iFld As DAO.Field
'Clr_Ays oAnFld
'Dim N%: N = 0
'For Each iFld In mDb.TableDefs(pNmt).Fields
'    Select Case iFld.Type
'    Case DAO.DataTypeEnum.dbText, DAO.DataTypeEnum.dbMemo
'        If iFld.Required Then ReDim Preserve oAnFld(N): oAnFld(N) = iFld.Name: N = N + 1
'    End Select
'Next
'Exit Function
'R: ss.R
'E: Fnd_AnFld_ReqTxt = True: ss.B cSub, cMod, "pNmt", pNmt
'End Function

'Function Fnd_AnFld_ReqTxt__Tst()
'Dim mDb As database: If Opn_Db_R(mDb, "p:\workingdir\MetaAll.mdb") Then Stop: GoTo E
'Dim mAnt$(): If SqlIntoAy(mAnt, "Select NmTbl from [$Tbl] where NmTbl=PKey", mDb) Then Stop: GoTo E
'Dim J%, N%: N = Siz_Ay(mAnt)
'For J = 0 To N - 1
'    Dim mAnFld$(): If Fnd_AnFld_ReqTxt(mAnFld, "$" & mAnt(J), mDb) Then Stop: GoTo E
'    Debug.Print mAnt(J) & ":" & ToStr_Ays(mAnFld)
'Next
'Exit Function
'E: Fnd_AnFld_ReqTxt_Tst = True
'End Function

'Function Fnd_Nm_ById(oNm$, pItm$, pId&) As Boolean
''Aim: Assume there is a table [${pItm}] having 2 fields: [{pItm}] & [Nm{pItm}].  Use [{pItm}] to find the name in table.
'Const cSub$ = "Fnd_Nm_ById"
'If Fnd_ValFmSql(oNm, Fmt_Str("Select Nm{0} from [${0}] where {0}={1}", pItm, pId)) Then ss.A 1, "pId not find in [${pItm}]": GoTo E
'Exit Function
'E: Fnd_Nm_ById = True: ss.B cSub, cMod, "pItm,pId", pItm, pId
'End Function
'Function Fnd_Id_ByNm(oId&, pItm$, pNm$) As Boolean
''Aim: Assume there is a table [${pItm}] having 2 fields: [{pItm}] & [Nm{pItm}].  Use [Nm{pItm}] to find the Id in table.
'Const cSub$ = "Fnd_Id_ByNm"
'If Fnd_ValFmSql(oId, Fmt_Str("Select {0} from [${0}] where Nm{0}='{1}'", pItm, pNm)) Then ss.A 1, "pNm not find in [${pItm}]": GoTo E
'Exit Function
'E: Fnd_Id_ByNm = True: ss.B cSub, cMod, "pItm,pNm", pItm, pNm
'End Function
'Function Fnd_AyRoot(oAyRoot&(), pNmt$, pPar$, pChd$) As Boolean
''Aim: In {pNmt} fields {pPar} & {pChr} are parent & child relation.  It is to find all Id of root to {oAyTblRoot}
'Const cSub$ = "Fnd_AyRoot"
'On Error GoTo R
'Dim mNmtChd$: mNmtChd$ = "[##CHD" & Format(Now, "YYYYMMDD HHMMSS") & "]"
'Dim mNmtPar$: mNmtPar$ = "[##PAR" & Format(Now, "YYYYMMDD HHMMSS") & "]"
'Dim mNmt$: mNmt = Q_S(pNmt, "[]")
'Dim mSql$
'mSql = Fmt_Str("Select Distinct {0} into {1} from {2}", pPar, mNmtPar, mNmt): If Run_Sql(mSql) Then ss.A 1: GoTo E
'mSql = Fmt_Str("Select Distinct {0} into {1} from {2}", pChd, mNmtChd, mNmt): If Run_Sql(mSql) Then ss.A 2: GoTo E
'mSql = Fmt_Str_ByLpAp("Select {pPar} from {mNmtPar} p left join {mNmtChd} c on p.{pPar}=c.{pChd} where c.{pChd} is null order by {pPar}" _
'    , "pPar,pChd,mNmtPar,mNmtChd,mNmt", pPar, pChd, mNmtPar, mNmtChd, mNmt)
'Fnd_AyRoot = SqlIntoAy(oAyRoot, mSql)
'Run_Sql "Drop Table" & mNmtChd
'Run_Sql "Drop Table" & mNmtPar
'Exit Function
'R: ss.R
'E: Fnd_AyRoot = True: ss.B cSub, cMod, "pNmt,pPar,pChd", pNmt, pPar, pChd
'End Function

'Function Fnd_AyRoot__Tst()
'If TblCrt_FmLnkLnt("p:\workingdir\MetaAll.mdb", "$TblR,$Tbl") Then Stop: GoTo E
'Dim mAyRoot&(): If Fnd_AyRoot(mAyRoot, "$TblR", "Tbl", "TblTo") Then Stop: GoTo E
'Dim mAnt$(): If SqlIntoAy(mAnt, "Select NmTbl from [$Tbl] where Tbl in (" & ToStr_AyLng(mAyRoot) & ")") Then Stop: GoTo E
'Debug.Print ToStr_Ays(mAnt)
'Debug.Print ToStr_AyLng(mAyRoot)
'Exit Function
'E: Fnd_AyRoot_Tst = True
'End Function

'Function Fnd_AyChd_ByRoot(oAyId&(), pNmt$, pRoot&, pPar$, pChd$ _
'    , Optional pRootFirst As Boolean = False _
'    , Optional pKeepAyId As Boolean = False) As Boolean
''Aim: In {pNmt} fields {pPar} & {pChr} are parent & child relation.  It is to find all Id of the tree of {pRoot} into {oAyId}
'Const cSub$ = "Fnd_AyChd_ByRoot"
'On Error GoTo R
'If Not pKeepAyId Then Clr_AyLng oAyId
'On Error GoTo R
'Dim N%, J
'If pRootFirst Then If Add_AyEleLng(oAyId, pRoot, True) Then Exit Function
'Dim mSql$: mSql = Fmt_Str_ByLpAp("Select {pChd} from {pNmt} where {pPar}={pRoot}", "pChd,pNmt,pPar,pRoot", pChd, Q_S(pNmt, "[]"), pPar, pRoot)
'Dim mAyId&(): If SqlIntoAy(mAyId, mSql) Then ss.A 2: GoTo E
'For J = 0 To Siz_Ay(mAyId) - 1
'    If Fnd_AyChd_ByRoot(oAyId, pNmt, mAyId(J), pPar, pChd, pRootFirst, True) Then ss.A 3: GoTo E
'Next
'If Not pRootFirst Then If Add_AyEleLng(oAyId, pRoot, True) Then Exit Function
'Exit Function
'R: ss.R
'E: Fnd_AyChd_ByRoot = True: ss.B cSub, cMod, "pNmt,pRoot,pPar,pChd,pRootFirst,pKeepAyId", pNmt, pRoot, pPar, pChd, pRootFirst, pKeepAyId
'End Function

'Function Fnd_AyChd_ByRoot__Tst()
'If TblCrt_FmLnkLnt("p:\workingdir\MetaAll.mdb", "$TblR,$Tbl") Then Stop: GoTo E
'Dim mAyRoot&(): If Fnd_AyRoot(mAyRoot, "$TblR", "Tbl", "TblTo") Then Stop: GoTo E
'Dim J%
'For J = 0 To Siz_Ay(mAyRoot) - 1
'    Dim mNmt$
'    If Fnd_ValFmSql(mNmt, "Select NmTbl from [$Tbl] where Tbl=" & mAyRoot(J)) Then Stop: GoTo E
'    Debug.Print mAyRoot(J); ": "; mNmt
'    Dim mAyId&(): If Fnd_AyChd_ByRoot(mAyId, "$TblR", mAyRoot(J), "Tbl", "TblTo", True) Then Stop: GoTo E
'    Dim mAnt$(): mAnt = SqlSy("Select NmTbl from [$Tbl] where Tbl in (" & ToStr_AyLng(mAyId) & ")") Then Stop: GoTo E
'    Debug.Print ToStr_Ays(mAnt)
'    Debug.Print ToStr_AyLng(mAyId)
'Next
'Exit Function
'E: Fnd_AyChd_ByRoot_Tst = True
'End Function

''Function Fnd_V(oV$, pV) As Boolean
''If VarType(pV) = vbString Then oV = oV: Exit Function
''Fnd_V = True
''End Function
'Function Fnd_Cmd(oCmd$, oRno&, pWs As Worksheet, pTar As Range) As Boolean
''Aim: Find the CtCommand of current pTar cell
'Const cSub$ = "Fnd_Cmd"
'If pTar.Count <> 1 Then GoTo E
'oRno = pTar.Row: 'If oRno < g.cRnoDta Then GoTo E
'Dim mCno%: mCno = pTar.Column
'If pTar.Interior.Color <> G.cColrCmd Then GoTo E
'Dim mV
'If pWs.Cells(2, mCno).Interior.Color <> G.cColrCmd Then GoTo E
'mV = pWs.Cells(2, mCno).Value
'If VarType(mV) <> vbString Then GoTo E
'If mV = "Cmd" Then mV = pTar.Value: If VarType(mV) <> vbString Then GoTo E
'oCmd = Replace(Replace(Replace(mV, vbCr, ""), vbLf, ""), " ", "")
'Exit Function
'E: Fnd_Cmd = True
'End Function
'Function Fnd_RowAyVal(oAyVal$(), pAyRno&(), pWs As Worksheet, pRno&, pNm$) As Boolean
''Aim: Find {oAyVal} at {AyRno} of {pNm}.  Assume there is a name of x{pNm} defining each column
'Const cSub = "RowAyVal"
'On Error GoTo R
'Clr_Ays oAyVal
'Dim mNm As Excel.Name
'If Fnd_Nm(mNm, pWs, "x" & pNm) Then ss.A 1: GoTo E
'Dim mRge As Range: Set mRge = mNm.RefersToRange
'Dim mCno As Byte: mCno = mRge.Column
'Dim J%
'For J = 0 To Siz_Ay(pAyRno) - 1
'    Add_AyEle oAyVal, Nz(pWs.Cells(pAyRno(J), mCno).Value, "")
'Next
'GoTo X
'R: ss.R
'E: Fnd_RowAyVal = True: ss.B cSub, cMod, "AyRno,pWs,pRno,pNm", ToStr_AyLng(pAyRno), ToStr_Ws(pWs), pRno, pNm
'X:
'End Function
'Function Fnd_Cno_ByNm(oCno As Byte, pWs As Worksheet, pNm$) As Boolean
'Const cSub$ = "Fnd_Cno_ByNm"
'Dim mNm As Excel.Name
'If Fnd_Nm(mNm, pWs, "x" & pNm) Then ss.A 1: GoTo E
'oCno = mNm.RefersToRange.Column
'Exit Function
'R: ss.R
'E: Fnd_Cno_ByNm = True: ss.B cSub, cMod, "pWs,pNm", ToStr_Ws(pWs), pNm
'End Function
'Function Fnd_RowVal(oRowVal$, pWs As Worksheet, pRno&, pLn$) As Boolean
''Aim: Find {oRowVal} at {pRno} of list name in {pLn}.  Assume there is names of xXXX defining each column
'Const cSub = "RowVal"
'On Error GoTo R
'oRowVal = ""
'Dim mAn$(): mAn = Split(pLn, ",")
'Dim J%
'For J = 0 To Siz_Ay(mAn) - 1
'    Dim mCno As Byte: If Fnd_Cno_ByNm(mCno, pWs, mAn(J)) Then ss.A 1: GoTo E
'    Dim mV: mV = pWs.Cells(pRno, mCno).Value
'    oRowVal = Add_Str(oRowVal, CStr(mV))
'Next
'Exit Function
'R: ss.R
'E: Fnd_RowVal = True: ss.B cSub, cMod, "pWs,pRno,pLn", ToStr_Ws(pWs), pRno, pLn
'End Function

'Function Fnd_RowVal__Tst()
'Dim mRowVal$: If Fnd_RowVal(mRowVal, Worksheets("Tbl"), 10, "NmTbl,NmTy1xxx") Then Stop: GoTo E
'Debug.Print mRowVal
'Exit Function
'E: Fnd_RowVal_Tst = True
'End Function

'Function Fnd_AyFfnRf(oAyFfnRf$(), pPrj As vbproject) As Boolean
''Aim: find {mAyFfnRf} of {pPrj}
'Const cSub$ = "Fnd_AyFfnRf"
'On Error GoTo R
'ReDim oAyFfnRf(pPrj.References.Count - 1)
'Dim iRf As VBIDE.Reference
'Dim J%: J = 0
'For Each iRf In pPrj.References
'    oAyFfnRf(J) = iRf.FullPath: J = J + 1
'Next
'Exit Function
'R: ss.R
'E: Fnd_AyFfnRf = True: ss.B cSub, cMod, "pPrj", ToStr_Prj(pPrj)
'End Function
'Function Fnd_Ffn_ByNmPrj(oFfnPrj$, pNmPrj$) As Boolean
'Const cSub$ = "Fnd_Ffn_ByNmPrj"
'Dim mPrj As vbproject: If Fnd_Prj(mPrj, pNmPrj) Then ss.A 1: GoTo E
'oFfnPrj$ = mPrj.FileName
'Exit Function
'E: Fnd_Ffn_ByNmPrj = True: ss.B cSub, cMod, "pNmPrj", pNmPrj
'End Function

'Function Fnd_Ffn_ByNmPrj__Tst()
'Dim mFfnPrj$: If Fnd_Ffn_ByNmPrj(mFfnPrj, "jj") Then Stop: GoTo E
'Debug.Print mFfnPrj
'E: Fnd_Ffn_ByNmPrj_Tst = True
'End Function

'Function Fnd_An_BySetNm_Sql(oAn$(), pSetNm$, Sql$) As Boolean
''Aim: Find {oAn} by setting all first field of {Sql} if it like pSetNm$
'Const cSub$ = "Fnd_An_BySetNm_Sql"
'Clr_Ays oAn
'Dim mAyLik$(): If Brk_Ln2Ay(mAyLik, pSetNm) Then ss.A 1: GoTo E
'Dim mRs As DAO.Recordset: If Opn_Rs(mRs, Sql) Then ss.A 2: GoTo E
'With mRs
'    While Not .EOF
'        Dim mV$: mV = .Fields(0).Value
'        If IsLikAyLik(mV, mAyLik) Then Add_AyEle oAn, mV
'        .MoveNext
'    Wend
'End With
'GoTo X
'R: ss.R
'E: Fnd_An_BySetNm_Sql = True: ss.B cSub, cMod, "pSetNm,Sql", pSetNm, Sql
'X: RsCls mRs
'End Function

'Function Fnd_An_BySetNm_Sql__Tst()
'Dim mAn$(), mSetNm$, mSql$
'mSql = "Select Distinct NmTbl from [$Tbl]"
'mSetNm = "Typ*"
'If Fnd_An_BySetNm_Sql(mAn, mSetNm, mSql) Then Stop
'Debug.Print Join(mAn, vbLf)
'End Function

'Function Fnd_Ws(oWs As Worksheet, pWb As Workbook, pNmWs$, Optional pSilent As Boolean) As Boolean
'Const cSub$ = "Fnd_Ws"
'On Error GoTo R
'Set oWs = pWb.Sheets(pNmWs)
'Exit Function
'R: ss.R
'E: Fnd_Ws = True: If Not pSilent Then ss.B cSub, cMod, "pWb,pNmWs", ToStr_Wb(pWb), pNmWs
'End Function
'Function Fnd_RnoLas(oRnoLas&, Rg As Range) As Boolean
''Aim: find first empty cell of a column {pCno} in {pWs} starting {pRnoFm}into {oRnoLas}
'Const cSub$ = "Fnd_RnoLas"
'On Error GoTo R
'If IsEmpty(Rg(1, 1).Value) Then oRnoLas = Rg.Row - 1: Exit Function
'Dim mRge As Range: Set mRge = Rg(1, 1)
'oRnoLas = mRge.End(xlDown).Row
'Exit Function
'R: ss.R
'E: Fnd_RnoLas = True: ss.B cSub, cMod, "Rg", ToStr_Rge(Rg)
'End Function
'Function Fnd_Ant_BySetNmt(oAnt$(), pSetNmt$, Optional pDb As database) As Boolean
''Aim: Find {oAnt} by {pSetNmt} in {pDb}
'Const cSub$ = "Fnd_Ant_BySetNmt"
'Clr_Ays oAnt
'Dim mAyLikNmt$(): mAyLikNmt = Split(pSetNmt$, CtComma)
'Dim J%
'For J = 0 To Siz_Ay(mAyLikNmt) - 1
'    Dim mAnt$(): If Fnd_Ant_ByLik(mAnt, Trim(mAyLikNmt(J)), pDb) Then ss.A 1: GoTo E
'    If Add_AyAtEnd(oAnt, mAnt) Then ss.A 1: GoTo E
'Next
'Exit Function
'R: ss.R
'E: Fnd_Ant_BySetNmt = True: ss.B cSub, cMod, "pSetNmt,pDb", pSetNmt, ToStr_Db(pDb)
'End Function

'Function Fnd_Ant_BySetNmt__Tst()
'Const cSub$ = "Fnd_Ant_BySetNmt_Tst"
'Dim mSetNmt$: mSetNmt = "mst*,tbl*"
'Dim mAntq$(): If Fnd_Ant_BySetNmt(mAntq, mSetNmt) Then Stop
'Shw_Dbg cSub, cMod, "mSetNmt,Result(mAntq)", mSetNmt, ToStr_Ays(mAntq)
'End Function

'Function Fnd_Antq_BySetNmtq(oAntq$(), pSetNmtq$, Optional pDb As database, Optional pQ$ = "") As Boolean
''Aim: Find {oAntq} by {pSetNmtq} in {pDb}
'Const cSub$ = "Fnd_Antq_BySetNmtq"
'Clr_Ays oAntq
'Dim mAyLikNmtq$(): mAyLikNmtq = Split(pSetNmtq$, CtComma)
'Dim J%
'
'For J = 0 To Siz_Ay(mAyLikNmtq) - 1
'    Dim mAntq$(): If Fnd_Antq_ByLik(mAntq, Trim(mAyLikNmtq(J)), pDb, pQ) Then ss.A 1: GoTo E
'    If Add_AyAtEnd(oAntq, mAntq) Then ss.A 1: GoTo E
'Next
'Exit Function
'R: ss.R
'E: Fnd_Antq_BySetNmtq = True: ss.B cSub, cMod, "pSetNmtq,pDb,pQ", pSetNmtq, ToStr_Db(pDb), pQ
'End Function

'Function Fnd_Antq_BySetNmtq__Tst()
'Const cSub$ = "Fnd_Antq_BySetNmtq_Tst"
'Dim mSetNmtq$: mSetNmtq = "mst*,tbl*"
'Dim mAntq$(): If Fnd_Antq_BySetNmtq(mAntq, mSetNmtq) Then Stop
'Shw_Dbg cSub, cMod, "mSetNmtq,Result(mAntq)", mSetNmtq, ToStr_Ays(mAntq)
'End Function

'Function Fnd_AnTxtSpec(oAnTxtSpec$(), Optional pDb As database) As Boolean
'Const cSub$ = "Fnd_AnTxtSpec"
'Dim mDb As database: Set mDb = DbNz(pDb)
'If Fnd_LoAyV_FmSql_InDb(mDb, "Select SpecName from MSysIMEXSpecs", "SpecName", oAnTxtSpec) Then
'    Dim mA$(): oAnTxtSpec = mA
'    ss.A 1: GoTo E
'End If
'Exit Function
'R: ss.R
'E: Fnd_AnTxtSpec = True: ss.B cSub, cMod, "pDb", ToStr_Db(pDb)
'End Function
'Function Fnd_AnTxtSpec__Tst()
'Dim mAnTxtSpec$(): If Fnd_AnTxtSpec(mAnTxtSpec) Then Stop
'Debug.Print ToStr_Ays(mAnTxtSpec)
'End Function
'Function Fnd_TxtSpecId(oTxtSpecId&, pNmSpec$, Optional pDb As database) As Boolean
'Const cSub$ = "Fnd_TxtSpecId"
'Set_Silent
'If Fnd_ValFmSql(oTxtSpecId, "Select SpecId from MSysIMEXSpecs where SpecName='" & pNmSpec & CtSngQ, pDb) Then GoTo E
'GoTo X
'E: Fnd_TxtSpecId = True
'X: Set_Silent_Rst
'End Function
'Function Fnd_AyDQry(oAyDQry() As d_Qry, QryNmLik$, Optional pInclQDpd As Boolean = False, Optional pAcs As Access.Application = Nothing) As Boolean
''Aim: Find the {AyDQry} of {QryNms} in {pDb}
'Const cSub$ = "Fnd_AyDQry"
'Dim mAcs As Access.Application: Set mAcs = Cv_Acs(pAcs)
'Dim mDb As database:        Set mDb = mAcs.CurrentDb
'
'Dim mAnq$(): If Fnd_Anq_ByLik(mAnq, QryNmLik, mDb) Then ss.A 1: GoTo E
'Dim N%: N = Siz_Ay(mAnq)
'If N = 0 Then
'    Dim mAyDQry() As d_Qry: oAyDQry = mAyDQry: Exit Function
'    Exit Function
'End If
'
'Dim mLasMaj%: mLasMaj = -1
'Dim J%, I%: I = 0
'For J = 0 To N - 1
'    Dim iQry As DAO.QueryDef: Set iQry = mDb.QueryDefs(mAnq(J))
'
'    ReDim Preserve oAyDQry(I)
'    Set oAyDQry(I) = New d_Qry
'    With oAyDQry(I)
'        If .Brk_Nmqs(iQry.Name) Then ss.xx 1, cSub, cMod: GoTo Nxt
'        .Typ = iQry.Type
'        On Error Resume Next
'        .Des = iQry.Properties("Description").Value
'        On Error GoTo 0
'        .Sql = iQry.Sql
'        If mLasMaj <> .Maj Then
'            If .Min <> 0 Then ss.A 1, "There is no Min Step 0", , "iQry.Name,QryNmLik,mLasMaj,mMaj", iQry.Name, QryNmLik, mLasMaj, .Maj: GoTo Nxt
'            If .Typ <> DAO.QueryDefTypeEnum.dbQSelect Then ss.A 1, "The query of minor step 0 must be select query", , "The Query,Query Type(DAO.QueryDefTypeEnum)", iQry.Name, iQry.Type: GoTo Nxt
'            mLasMaj = .Maj
'        End If
'    End With
'    I = I + 1
'Nxt:
'Next
'If pInclQDpd Then
'    For J = 0 To N - 1
'        With oAyDQry(J)
'            .LnTbl = ToStr_SqlLnt(.Sql)
'        End With
'    Next
'End If
'GoTo X
'R: ss.R
'E: Fnd_AyDQry = True: ss.B cSub, cMod, "QryNmLik,pInclQDpd,pAcs", QryNmLik, pInclQDpd, ToStr_Acs(pAcs)
'X:  Set mDb = Nothing
'End Function

'Function Fnd_AyDQry__Tst()
'Const cFfnCsv$ = "c:\aa.csv"
'Const cDir$ = "D:\SPLHalfWayHouse\MPSDetail\VerNew@2007_01_04\WorkingDir\PgmObj\"
'
'Dim mAyFb$(): If Fnd_AyFn(mAyFb, cDir, "*.mdb") Then Stop
'Dim I%
'
'Dim mF As Byte: If Opn_Fil_ForOutput(mF, cFfnCsv, True) Then Stop
'Write #mF, "Mdb";
'Dim mDQry As New d_Qry
'If mDQry.WrtHdr(mF) Then Stop
'
'For I = 0 To Siz_Ay(mAyFb) - 1
'    Dim mNmQs$: mNmQs = ""
'    Select Case mAyFb(I)
'    Case "MPS_GenDta.mdb":    mNmQs = "qryMPS"
'    Case "MPS_GenRpt.mdb":    mNmQs = "qryMPS"
'    Case "MPS_Odbc.mdb":      mNmQs = "qryOdbcMPS"
'    Case "MPS_RfhCusGp.mdb":  mNmQs = "qryRfhCusGp"
'    Case "RfhFc.mdb":     mNmQs = "qryFc,qryOdbcFc"
'    End Select
'    If Len(mNmQs) = 0 Then GoTo Nxt
'
'    Dim mAnQs$(): mAnQs = Split(mNmQs, CtComma)
'    Dim mDb As database:
'    Do
'        If Opn_Db(mDb, cDir & mAyFb(I), True) Then Stop
'        Dim N%: N = Siz_Ay(mAnQs)
'
'        Dim J%
'        For J = 0 To N - 1
'            Dim mAyDQry() As d_Qry: If Fnd_AyDQry(mAyDQry, mAnQs(J) & "*", True, mDb) Then Stop
'            Dim K%
'            For K = 0 To Siz_AyDQry(mAyDQry) - 1
'                If mAyDQry(K).Wrt(mF, mAyFb(I)) Then Stop
'            Next
'        Next
'    Loop Until True
'    mDb.Close
'Nxt:
'Next
'Close #mF
'Dim mWb As Workbook: If Opn_Wb_RW(mWb, cFfnCsv) Then Stop
'mWb.Application.Visible = True
'End Function

'Function Fnd_MaxLin%(pLines$)
'Dim J%, L%, mAys$()
'mAys = Split(pLines, vbLf)
'For J = 0 To Siz_Ay(mAys) - 1
'    If L < Len(mAys(J)) Then L = Len(mAys(J))
'Next
'Fnd_MaxLin = L
'End Function
'Function Fnd_LnFld_ByNmtq(oLnFld$, Qry_or_Tbl_Nm$, Optional pDb As database, Optional pInclTypFld As Boolean = False) As Boolean
''Aim: Find {oLnFld} by {Qry_or_Tbl_Nm} in {pDb} to return if {pInclTypFld} & with {pSepChr}
'Const cSub$ = "Fnd_LnFld_ByNmtq"
'On Error GoTo R
'Dim mDb As database: Set mDb = DbNz(pDb)
'If IsTbl(Qry_or_Tbl_Nm, mDb) Then oLnFld = ToStr_Flds(mDb.TableDefs(Qry_or_Tbl_Nm).Fields, pInclTypFld): Exit Function
'If IsQry(Qry_or_Tbl_Nm, mDb) Then oLnFld = ToStr_Flds(mDb.QueryDefs(Qry_or_Tbl_Nm).Fields, pInclTypFld): Exit Function
'ss.A 1, "Given Qry_or_Tbl_Nm not exist in pDb"
'GoTo E
'R: ss.R
'E: Fnd_LnFld_ByNmtq = True: ss.B cSub, cMod, "Qry_or_Tbl_Nm,pDb", Qry_or_Tbl_Nm, ToStr_Db(pDb)
'End Function
'Function Fnd_LnFld_ByNmq(oLnFld$, QryNm$, Optional pInclTypFld As Boolean = False, Optional pSepChr$ = CtComma, Optional pDb As database) As Boolean
''Aim: Find {oLnFld} by {QryNm} in {pDb} to return if {pInclTypFld} & with {pSepChr}
''       Always look for any fields in {oLnFld} begins with yymd_.
''       If so, replace the field by:
''           Cdate(IIf(yymd_xxx=0,0,IIf(yymd_xxx=99999999,'9999/12/31',format(yymd_xxx,'0000\/00\/00')))) as xxx
''       Else
''           return oLnFld as "*"
'Const cSub$ = "Fnd_LnFld_ByNmq"
'On Error GoTo R
'Dim mDb As database: Set mDb = DbNz(pDb)
'oLnFld = ToStr_Flds(mDb.QueryDefs(QryNm).Fields, pInclTypFld, , pSepChr)
'If Left(oLnFld, 3) = "Err" Then ss.A 1, "Cannot obtain field list from query", , "QryNm,Sql,oLnFld", QryNm, ToSql_Nmq(QryNm), oLnFld: GoTo E
'oLnFld = Cv_LnFld(oLnFld)
'Exit Function
'R: ss.R
'E: Fnd_LnFld_ByNmq = True: ss.B cSub, cMod, "QryNm,pInclTypFld,pSepChr,pDb", QryNm, pInclTypFld, pSepChr, ToStr_Db(pDb)
'End Function
'Function Fnd_LnFld_ByNmt(oLnFld$, pNmt$, Optional pInclTypFld As Boolean = False, Optional pSepChr$ = CtComma, Optional pDb As database) As Boolean
'Const cSub$ = "Fnd_LnFld_ByNmt"
'On Error GoTo R
'Dim mDb As database: Set mDb = DbNz(pDb)
'oLnFld = Cv_LnFld(ToStr_Flds(mDb.TableDefs(Rmv_SqBkt(pNmt)).Fields, pInclTypFld, , pSepChr))
'Exit Function
'R: ss.R
'E: Fnd_LnFld_ByNmt = True: ss.B cSub, cMod, "pNmt,pInclTypFld,pSepChr", pNmt, pInclTypFld, pSepChr
'End Function

'Function Fnd_LnFld_ByNmt__Tst()
'Const cSub$ = "Fnd_LnFld_ByNmt_Tst"
'Dim mLnFld$, mNmt$
'Dim mResult As Boolean
'Dim mCase As Byte: mCase = 1
'Select Case mCase
'Case 1: mNmt = "tmpFc_XlsKMR"
'End Select
'mResult = Fnd_LnFld_ByNmt(mLnFld, mNmt)
'Shw_Dbg cSub, cMod, , "mLnFld, mNmt", mLnFld, mNmt
'End Function

'Function Fnd_Sffn_LgcMdb(oFbLgc$, pNmLgc$) As Boolean
'Const cSub$ = "Fnd_Sffn_LgcMdb"
'If Not IsTbl("Av_LgcMdb") Then
'    If TblCrt_FmLnkNmt(Sdir_PgmObj & "Av.mdb", "Av_LgcMdb") Then ss.A 1: GoTo E
'End If
'If Fnd_ValFmSql(oFbLgc, "Select FbLgc from Av_LgcMdb where NmLgc='" & pNmLgc & CtSngQ) Then ss.A 2: GoTo E
'If Left(oFbLgc$, 2) = ".\" Then oFbLgc = Sdir_PgmObj & mID(oFbLgc$, 3)
'Exit Function
'R: ss.R
'E: Fnd_Sffn_LgcMdb = True: ss.B cSub, cMod, "pNmLgc"
'End Function

'Function Fnd_Sffn_LgcMdb__Tst()
'Dim mFbLgc$: If Fnd_Sffn_LgcMdb(mFbLgc, "AddTbl") Then Stop
'Debug.Print mFbLgc
'End Function

'Function Fnd_Sffn_LgcMdbTmp(oFbOldQsTmp$, pNmLgc$) As Boolean
'Const cSub$ = "Fnd_Sffn_LgcMdbTmp"
'Dim mFbLgc$: If Fnd_Sffn_LgcMdb(mFbLgc, pNmLgc$) Then ss.A 1: GoTo E
'oFbOldQsTmp = Sdir_TmpLgc & "tmp" & Nam_FilNam(mFbLgc$)
'Exit Function
'R: ss.R
'E: Fnd_Sffn_LgcMdbTmp = True: ss.B cSub, cMod, "pNmLgc"
'End Function

'Function Fnd_Sffn_LgcMdbTmp__Tst()
'Dim mFbOldQsTmp$: If Fnd_Sffn_LgcMdbTmp(mFbOldQsTmp, "AddTbl") Then Stop
'Debug.Print mFbOldQsTmp
'End Function

'Function Fnd_AnQs(oAnQs$(), Optional pLikQry$ = "qry*", Optional pDb As database) As Boolean
'Dim mDb As database: Set mDb = DbNz(pDb)
'Dim mNmQsLas$, mNmQsCur$, N%, iQry As DAO.QueryDef
'N = 0
'For Each iQry In mDb.QueryDefs
'    If Left(iQry.Name, 1) = "~" Then GoTo Nxt
'    If Not iQry.Name Like pLikQry Then GoTo Nxt
'    mNmQsCur = Fnd_NmQs(iQry.Name)
'    If mNmQsCur = "" Then GoTo Nxt
'    If mNmQsLas <> mNmQsCur Then
'        ReDim Preserve oAnQs(N)
'        oAnQs(N) = mNmQsCur
'        N = N + 1
'        mNmQsLas = mNmQsCur
'    End If
'Nxt:
'Next
'If N = 0 Then Clr_Ays oAnQs
'End Function

'Function Fnd_AnQs__Tst()
'Dim mAnQs$()
'If Fnd_AnQs(mAnQs) Then Stop
'Debug.Print Join(mAnQs, vbLf)
'End Function

'Function Fnd_AnDpd(oAnDpd$(), QryNm$, Optional pAcs As Access.Application = Nothing) As Boolean
''Aim: The {AnDpd} by {QryNm}
'Const cSub$ = "Fnd_AnDpd"
'Dim mAcs As Access.Application: Set mAcs = Cv_Acs(pAcs)
'Dim mCase As Byte: mCase = 2
'Select Case mCase
'Case 1
'    Dim mQry As Access.AccessObject: Set mQry = mAcs.CurrentData.AllQueries(QryNm)
'    Dim mDpdInfo As Access.DependencyInfo: Set mDpdInfo = mQry.GetDependencyInfo
'    Dim N%: N = mDpdInfo.Dependencies.Count
'    If N = 0 Then Clr_Ays oAnDpd: Exit Function
'    Dim J%
'    ReDim oAnDpd(N - 1)
'    On Error Resume Next
'    For J = 0 To N - 1
'        oAnDpd(J) = "?"
'        oAnDpd(J) = mDpdInfo.Dependencies(J).Name
'    Next
'Case 2
'    Dim mSql$: mSql = mAcs.CurrentDb.QueryDefs(QryNm).Sql
'    If Fnd_SqlTbl(oAnDpd, mSql) Then ss.A 1: GoTo E
'End Select
'Exit Function
'E: Fnd_AnDpd = True: ss.B cSub, cMod, "QryNm,pAcs", QryNm, ToStr_Acs(pAcs)
'End Function

'Function Fnd_AnDpd__Tst()
'Dim mAnq$(), mAnt$()
'If Fnd_Anq_ByLik(mAnq, "*") Then Stop: GoTo E
'Dim J%
'For J = 0 To Siz_Ay(mAnq) - 1
'    Debug.Print mAnq(J),
'    If Fnd_AnDpd(mAnt, mAnq(J)) Then Stop: GoTo E
'    Debug.Print ToStr_Ays(mAnt)
'Next
'Exit Function
'E: Fnd_AnDpd_Tst = True
'End Function

'Function Fnd_SqlTbl(oAnt$(), Sql$) As Boolean
'ReDim oAnt(0): oAnt(0) = Sql
'End Function
'Function Fnd_Idx(oIdx%, pAy$(), pV$) As Boolean
''Aim: Find {pV} in {pAy} by return {oIdx}
'Const cSub$ = "Fnd_Idx"
'On Error GoTo R
'For oIdx = 0 To Siz_Ay(pAy) - 1
'    If pAy(oIdx) = pV Then Exit Function
'Next
'oIdx = -1
'Exit Function
'R: ss.R
'E: Fnd_Idx = True: ss.B cSub, cMod, cSub, cMod, "pAy,pV", ToStr_Ays(pAy), pV
'End Function
'Function Fnd_IdxLng(oIdx%, pAyLng&(), pLng&) As Boolean
''Aim: Find {pLng} in {pAyLng} by return {oIdx}
'Const cSub$ = "Fnd_IdxLng"
'On Error GoTo R
'For oIdx = 0 To Siz_Ay(pAyLng) - 1
'    If pAyLng(oIdx) = pLng Then Exit Function
'Next
'oIdx = -1
'Exit Function
'R: ss.R
'E: Fnd_IdxLng = True: ss.B cSub, cMod, "pAyLng,pLng", ToStr_AyLng(pAyLng), pLng
'End Function
'Function Fnd_IdxByt(oIdx%, pAyByt() As Byte, pByt As Byte) As Boolean
''Aim: Find {pLng} in {pAyLng} by return {oIdx}
'Const cSub$ = "Fnd_IdxByt"
'On Error GoTo R
'For oIdx = 0 To Siz_Ay(pAyByt) - 1
'    If pAyByt(oIdx) = pByt Then Exit Function
'Next
'oIdx = -1
'Exit Function
'R: ss.R
'E: Fnd_IdxByt = True: ss.B cSub, cMod, "pAyByt,pByt", ToStr_AyByt(pAyByt), pByt
'End Function
'Function Fnd_FfnDtf$(TarFb$, TarTn$)
'Dim mDir$
'If TarFb = "" Then
'    mDir = Fct.CurMdbDir & "DTF\"
'Else
'    mDir = Fct.Nam_DirNam(TarFb) & "DTF\"
'End If
'Crt_Dir mDir
'Fnd_FfnDtf = mDir & TarTn & ".Dtf"
'End Function
'Function Fnd_Fn_By_Tp_n_CurFnn(oFn$, pTp$, pCurFnn$) As Boolean
'Const cSub$ = "Fnd_Fn_By_Tp_n_CurFn"
''Aim: Assume pCurFnn is in fmt of xxxxx_nnnn  It is to return xxxxx_<tp>_Dta.mdb
'Dim p%: p = InStrRev(pCurFnn, "_")
'oFn = Left(pCurFnn, p) & pTp & "_Dta.mdb"
'Exit Function
'Fnd_Fn_By_Tp_n_CurFnn = True
'End Function

'Function Fnd_Fn_By_Tp_n_CurFnn__Tst()
'Dim mFn$: If Fnd_Fn_By_Tp_n_CurFnn(mFn, "123", "xxxx_nnnn") Then Stop
'Debug.Print mFn
'End Function

'Function Fnd_MaxEle(oIdx%, pAy$()) As Boolean
'Dim N%: N = Siz_Ay(pAy)
'If N = 0 Then oIdx = -1: Exit Function
'Dim J%, mMax$
'oIdx = 0: mMax = pAy(0)
'For J = 1 To N - 1
'    If pAy(J) > mMax Then oIdx = J: mMax = pAy(J)
'Next
'End Function
'Function Fnd_MinEle(oIdx%, pAy$()) As Boolean
'Dim N%: N = Siz_Ay(pAy)
'If N = 0 Then oIdx = -1: Exit Function
'Dim J%, mMin$
'oIdx = 0: mMin = pAy(0)
'For J = 1 To N - 1
'    If pAy(J) < mMin Then oIdx = J: mMin = pAy(J)
'Next
'End Function
'Function Fnd_LoQVal_ByFrm(oLoQVal$, pFrm As Access.Form, pAnCtl$()) As Boolean
'Const cSub$ = "Fnd_LoQVal_ByFrm"
''Aim: Find {oLoQVal} by the control's NewValue in {pFrm} using {pAnCtl} as the control's name.
'Dim J%, N%: N = Siz_Ay(pAnCtl)
'oLoQVal = ""
'Dim mA$
'For J = 0 To N - 1
'    If Fnd_QVal_ByFrm(mA, pFrm, pAnCtl(J)) Then ss.A 1: GoTo E
'    oLoQVal = Add_Str(oLoQVal, mA)
'Next
'Exit Function
'R: ss.R
'E: Fnd_LoQVal_ByFrm = True: ss.B cSub, cMod, "Frm,pAnCtl", ToStr_Frm(pFrm), Join(pAnCtl, CtComma)
'End Function

'Function Fnd_LoQVal_Frm__Tst()
'Const cNmFrm$ = "frmIIC_Tst"
'Dim mFrm As Access.Form: If FrmOpn(cNmFrm, , , mFrm) Then Stop: GoTo E
'Dim mAn$(): mAn = Split("ItemClass,Des,ICGL,ICALP5", CtComma)
'Dim mLstQVal$: If Fnd_LoQVal_ByFrm(mLstQVal, mFrm, mAn) Then Stop: GoTo E
'Debug.Print mLstQVal
'Exit Function
'E: Fnd_LoQVal_Frm_Tst = True
'End Function

'Function Fnd_LoAsg_InFrm(oLoAsg$, pFrm As Access.Form, pLm$, Optional oLoChgd$) As Boolean
'Const cSub$ = "Fnd_LoAsg_InFrm"
''Aim: Find {oLoAsg} & {oLoChgd$} by the control's NewValue in {pFrm} using {pLm} as the control's name.
''     pLm     is fmt of aaa=xxx,bbb,ccc                            aaa,bbb,ccc will be used to list of control name. xxx,bbb,ccc will be used in {oLoAsg} & {oLoChgd}
''     oLoAsg  is fmt of xxx='nnnn',bbb=nnnn                     which is the [ssss] part of "Update tttt set ssss where wwww"
''     oLoChgd is fmt of xxx=[oooo]<--[nnnn]|bbb=[oooo]<--[nnnn] which will be show in status.
'Dim mAnCtl$(), mAnAsg$(): If Brk_Lm_To2Ay(mAnCtl, mAnAsg, pLm) Then ss.A 1: GoTo E
'Dim J%, N%: N = Siz_Ay(mAnCtl)
'oLoAsg = "": oLoChgd = ""
'Dim mA$, mB$
'For J = 0 To N - 1
'    If Fnd_Asg_InFrm(mA, pFrm, mAnCtl(J), mAnAsg(J), , mB) Then ss.A 2: GoTo E
'    If mA <> "" Then
'        oLoAsg = Add_Str(oLoAsg, mA)
'        oLoChgd = Add_Str(oLoChgd, mB, vbCrLf)
'    End If
'Next
'Exit Function
'R: ss.R
'E: Fnd_LoAsg_InFrm = True: ss.B cSub, cMod, cSub, cMod, "Frm,pLm", ToStr_Frm(pFrm), pLm$
'End Function

'Function Fnd_LoAsg_InFrm__Tst()
'Const cNmFrm$ = "frmIIC_Tst"
'Dim mFrm As Access.Form: If FrmOpn(cNmFrm, , , mFrm) Then Stop: GoTo E
'Dim mLm$: mLm = "ItemClass,Des=ICDES,ICGL,ICALP5"
'Dim mLoAsg$, mLoChgd$: If Fnd_LoAsg_InFrm(mLoAsg, mFrm, mLm, mLoChgd) Then Stop
'Debug.Print ToStr_NmV("mLoAsg", mLoAsg)
'Debug.Print ToStr_NmV("mLoChgd", mLoChgd)
'Exit Function
'E: Fnd_LoAsg_InFrm_Tst = True
'End Function

'Function Fnd_Asg_InFrm(oAsg$, pFrm As Access.Form, pNmCtl$, Optional pNmAsg$ = "", Optional pAlwNull As Boolean = False, Optional oChgd$) As Boolean
''Aim: return {oAsg},{oChgd} as ""                       if the Control of name {pNmCtl} in {pFrm} having equal .Value or .OldValue
''     else
''     return {oAsg} as mNmAsg=<QuotedValue>, and
''            {oChgd}as mNmAsg={OldVal}<--{NewVal} from the control {pNm} in {pFrm}.
''       Note:   If the new value is Null,
''                   if the field is num,str or bool, oAsg will return as mNmAsg=0 or '' or false
''                   other field type will return err.
'Const cSub$ = "Fnd_Asg_InFrm"
'On Error GoTo R
'oAsg = "": oChgd = ""
'Dim mNmAsg$: mNmAsg = NonBlank(pNmAsg, pNmCtl)
'Dim mCtl As Access.Control: If Fnd_Ctl(mCtl, pFrm, pNmCtl) Then ss.A 1: GoTo E
'Dim mVNew, mVOld: mVNew = mCtl.Value: mVOld = mCtl.OldValue
'If mVNew = mVOld Then Exit Function
'Dim mTypSim As eTypSim: mTypSim = VarType(mVNew)
'If mTypSim = vbNull Then
'    If pAlwNull Then
'        oAsg = mNmAsg & "=Null"
'        oChgd = mNmAsg & "=Null<--[" & mVOld & "]"
'        Exit Function
'    End If
'    Dim mRs As DAO.Recordset: Set mRs = pFrm.Recordset
'    mTypSim = DaoTyToSim(mRs.Fields(pNmCtl).Type)
'    Select Case mTypSim
'    Case eTypSim_Bool: oAsg = mNmAsg & "=False": oChgd = mNmAsg & "=[False]<--[" & mVOld & "]"
'    Case eTypSim_Num: oAsg = mNmAsg & "=0":      oChgd = mNmAsg & "=[0]<--[" & mVOld & "]"
'    Case eTypSim_Str: oAsg = mNmAsg & "=''":     oChgd = mNmAsg & "=[]<--[" & mVOld & "]"
'    Case Else
'        ss.A 1, "The control having a null and it is not Bool,Num or Str", , "pFrm,pNmCtl,mNmAsg,SimTyp", ToStr_Frm(pFrm), pNmCtl, mNmAsg, mTypSim
'        GoTo E
'    End Select
'    Exit Function
'End If
'mTypSim = VarToSimTy(mVNew)
'Select Case mTypSim
'Case eTypSim_Bool, eTypSim_Num, eTypSim_Str
'    oAsg = mNmAsg & "=" & Q_V(mVNew): oChgd = mNmAsg & "=[" & mVOld & "]<--[" & mVNew & "]"
'Case Else
'    ss.A 1, "The control having a value not being (Num,Bool,Str)", , "The Ctl's NewVal SimTyp", mTypSim
'    GoTo E
'End Select
'Exit Function
'R: ss.R
'E: Fnd_Asg_InFrm = True: ss.B cSub, cMod, "pFrm,pNmCtl,mNmAsg", ToStr_Frm(pFrm), pNmCtl, mNmAsg
'End Function
'Function Fnd_QVal_ByFrm(oQVal$, pFrm As Access.Form, pNmCtl$, Optional pAlwNull As Boolean = False) As Boolean
''Aim: Find {oQVal} as a quoted value for the control {pNmCtl}.Value in {pFrm}.
''     Only Num, Str or Bool type is allowed.
''     Null value will return 0, '' or False
'Const cSub$ = "Fnd_QVal_ByFrm"
'On Error GoTo R
'Dim mV: mV = pFrm.Controls(pNmCtl).Value
'Dim mTypSim As eTypSim
'
'If VarType(mV) = vbNull Then
'    If pAlwNull Then oQVal = "Null": Exit Function
'    Dim mRs As DAO.Recordset: Set mRs = pFrm.Recordset
'    mTypSim = DaoTyToSim(mRs.Fields(pNmCtl).Type)
'    Select Case mTypSim
'    Case eTypSim_Bool:  oQVal = "False"
'    Case eTypSim_Num:   oQVal = "0"
'    Case eTypSim_Str:   oQVal = "''"
'    Case Else:          ss.A 1, "The control having a null and it is not Bool,Num or Str": GoTo E
'    End Select
'    Exit Function
'End If
'mTypSim = VarToSimTy(mV)
'Select Case mTypSim
'Case eTypSim_Bool, eTypSim_Num, eTypSim_Str, eTypSim_Dte
'    oQVal = Q_V(mV)
'Case Else
'    ss.A 2, "The control having a value not in (Num,Bool,Str,Dte)": GoTo E
'End Select
'Exit Function
'R: ss.R
'E: Fnd_QVal_ByFrm = True: ss.B cSub, cMod, "pFrm,pNmCtl,SimTyp of the ctl.value", ToStr_Frm(pFrm), pNmCtl, mTypSim
'End Function
'Function Fnd_AyMacroStr_InStr(oAyMacroStr$(), pInStr$) As Boolean
'Dim mP%, mA%, mB%, J%, mN As Byte
'Clr_Ays oAyMacroStr
'mP = 1
'Do
'    mA = InStr(mP, pInStr, "{")
'    If mA <= 0 Then Exit Function
'    mB = InStr(mP + 1, pInStr, "}")
'    If mB <= 0 Then Exit Function
'    Dim mMacro$: mMacro = mID(pInStr, mA, mB - mA + 1)
'    Dim mFnd As Boolean: mFnd = False
'    For J = 0 To mN - 1
'        If oAyMacroStr(J) = mMacro Then mFnd = True: Exit For
'    Next
'    If Not mFnd Then
'        ReDim Preserve oAyMacroStr(mN)
'        oAyMacroStr(mN) = mMacro: mN = mN + 1
'    End If
'    mP = mB + 1
'Loop
'End Function

'Function Fnd_AyMacroStr_InStr__Tst()
'Dim mDtfTp$: If Fnd_ResStr(mDtfTp, "DtfTp", True) Then Stop
'Dim mAyMacroStr$(): If Fnd_AyMacroStr_InStr(mAyMacroStr, mDtfTp) Then Stop
'Debug.Print Join(mAyMacroStr, vbLf)
'End Function

'Function Fnd_Ffn_Fm_LnkXlsNmt(oFx$, pLnkXlsNmt$) As Boolean
''Aim: Find {oFx} from {pLnkXlsNmt} which is table name of a linked Excel
'Const cSub$ = "Fnd_Ffn_Fm_LnkXlsNmt"
'On Error GoTo R
'Dim mCnn$: mCnn = CurrentDb.TableDefs(pLnkXlsNmt).Connect
'If Left(mCnn, 10) <> "Excel 8.0;" Then ss.A 1, "Given pLnkXlsNmt does not have connection string starts with [Excel 8.0;]", , "pLnkXlsNmt,CnnStr", pLnkXlsNmt, mCnn: GoTo E
'Dim mP%: mP = InStr(mCnn, "DATABASE="): If mP <= 0 Then ss.A 1, "Given pLnkXlsNmt connection string should be [DATABASE=]", , "pLnkXlsNmt,CnnStr", pLnkXlsNmt, mCnn: GoTo E
'oFx = mID(mCnn, mP + 9)
'Exit Function
'R: ss.R
'E: Fnd_Ffn_Fm_LnkXlsNmt = True: ss.B cSub, cMod, cSub, cMod
'End Function

'Function Fnd_Ffn_Fm_LnkXlsNmt__Tst()
'Const cSub$ = "Fnd_Ffn_Fm_LnkXlsNmt"
'Const cFfn$ = "c:\Book1.xls"
'Dim mWb As Workbook: If Crt_Wb(mWb, cFfn, True) Then Stop
'Cls_Wb mWb, True
'If TblCrt_FmLnkXls(cFfn) Then Stop
'Dim mFfn$, mA$
'mA = "Sheet1": If Fnd_Ffn_Fm_LnkXlsNmt(mFfn, mA) Then Stop Else Debug.Print mA, mFfn
'mA = "Sheet2": If Fnd_Ffn_Fm_LnkXlsNmt(mFfn, mA) Then Stop Else Debug.Print mA, mFfn
'mA = "Sheet3": If Fnd_Ffn_Fm_LnkXlsNmt(mFfn, mA) Then Stop Else Debug.Print mA, mFfn
'End Function

'Function Fnd_Prm_FmTblPrm(oTrc&, oNmLgc$, Optional oLm$) As Boolean
''Aim: Find {oLn} & {oAyV} from tblPrm
''     Assume tblPrm has only 1 rec and is:Trc,NmLgc,Lm
'Const cSub$ = "Fnd_Prm_FmTblPrm"
'On Error GoTo R
'With CurrentDb.OpenRecordset("Select * from tblPrm")
'    oTrc = !Trc
'    oNmLgc = !NmLgc
'    oLm = Nz(!Lm, "")
'    .Close
'End With
'Exit Function
'R: ss.R
'E: Fnd_Prm_FmTblPrm = True: ss.B cSub, cMod
'End Function
'Function Fnd_PrcDcl(oPrcDcl$, pMod$, pNmPrc$) As Boolean
''Aim: Get the 'Aim' lines into {oPrcDcl}.  Aim lines: first 50 lines with first line start with 'Aim and subsequent lines begin with '
'Const cSub$ = "Fnd_PrcDcl"
'On Error GoTo R
'Const cMaxLen% = 250
'Dim mS$
'Set_Silent
'If Fnd_PrcBody(mS, pMod, pNmPrc, , True) Then
'    If Fnd_PrcBody(mS, pNmPrc, pMod, , True) Then ss.A 1: GoTo E
'End If
'Dim mAy$(): mAy = Split(mS, vbCrLf)
'
''Put the Function Fnd_declaration lines into mAy() first
'oPrcDcl = ""
'Dim J%, I%, N%: N% = Fct.MinInt(50, Siz_Ay(mAy) - 1)
'For J = 0 To N
'    oPrcDcl = Add_Str(oPrcDcl, mAy(J), vbLf)
'    If Right(mAy(J), 1) <> "_" Then Exit For
'Next
''Find 'Aim
'For J = J To N
'    If Left(mAy(J), 4) = "'Aim" Then
'        For I = J To N
'            If Left(mAy(I), 1) <> CtSngQ Then GoTo X
'            oPrcDcl = oPrcDcl & vbCrLf & mAy(I)
'        Next
'    End If
'Next
'GoTo X
'R: ss.R
'E: Fnd_PrcDcl = True: ss.C cSub, cMod, "pMod,pNmPrc", pMod, pNmPrc
'X: Set_Silent_Rst
'End Function
'Function Fnd_PrcDcl__Tst()
'Const cSub$ = "Fnd_PrcDcl_Tst"
'Dim mPrcDcl$, mNmPrj_Nmm$, mNmPrc$
'Dim mRslt As Boolean, mCase As Byte
'mCase = 2
'Select Case mCase
'Case 1
'    mNmPrj_Nmm = cLib & ".Fnd"
'    mNmPrc = "PrcDcl"
'Case 2
'    mNmPrj_Nmm = cLib & ".Gen"
'    mNmPrc = "Doc"
'Case 3
'    mNmPrj_Nmm = cLib & ".Read"
'    mNmPrc = "Def_FmtTbl"
'Case 4
'    mNmPrj_Nmm = cLib & ".Bld"
'    mNmPrc = "OdbcQs_ByAySelSql_ByDsn"
'End Select
'mRslt = Fnd_PrcDcl(mPrcDcl, mNmPrj_Nmm, mNmPrc)
'Shw_DbgWin
'Debug.Print mPrcDcl
'End Function
'Function Fnd_AyCno_XInRow(pWs As Worksheet, pRno&, pCnoFm As Byte, pCnoTo As Byte) As Byte()
'Dim iCno As Byte, AyCno() As Byte, nCol As Byte
'For iCno = pCnoFm To pCnoTo
'    If pWs.Cells(pRno, iCno).Value = "X" Then
'        nCol = nCol + 1
'        ReDim Preserve AyCno(0 To nCol - 1)
'        AyCno(nCol - 1) = iCno
'    End If
'Next
'Fnd_AyCno_XInRow = AyCno()
'End Function
'Function Fnd_AyDir(oAyDir$(), pDir$) As Boolean
'Const cSub$ = "Fnd_AyDir"
''History: Created on=2006/08/15; Modified on=2006/08/15
''Aim: Get a list of sub-dir in an Array (Start Index is 1) of a dir {pDir}
''==Start
'If Not IsDir(pDir) Then ss.A 1: GoTo E
'Dim mSubDir$, AyLst$(), N As Byte
'mSubDir = VBA.Dir(pDir & "*.*", vbDirectory)
'While mSubDir <> ""
'    If mSubDir <> "." And mSubDir <> ".." Then
'        If GetAttr(pDir & mSubDir) And vbDirectory Then
'            ReDim Preserve AyLst(0 To N)
'            AyLst(N) = mSubDir
'            N = N + 1
'        End If
'    End If
'    mSubDir = VBA.Dir
'Wend
'oAyDir = AyLst
'Exit Function
'R: ss.R
'E: Fnd_AyDir = True: ss.B cSub, cMod, "pDir", pDir
'End Function
'Function Fnd_AyFld(oAnFld$(), Qry_or_Tbl_Nm$) As Boolean
'Const cSub$ = "Fnd_AyFld"
'Dim J As Byte
'If IsTbl(Qry_or_Tbl_Nm) Then
'    ReDim oAnFld(0 To CurrentDb.TableDefs(Qry_or_Tbl_Nm).Fields.Count - 1)
'    For J = 0 To CurrentDb.TableDefs(Qry_or_Tbl_Nm).Fields.Count - 1
'        oAnFld(J) = CurrentDb.TableDefs(Qry_or_Tbl_Nm).Fields(J).Name
'    Next
'    Exit Function
'End If
'If IsQry(Qry_or_Tbl_Nm) Then
'    ReDim oAnFld(0 To CurrentDb.QueryDefs(Qry_or_Tbl_Nm).Fields.Count - 1)
'    For J = 0 To CurrentDb.QueryDefs(Qry_or_Tbl_Nm).Fields.Count - 1
'        oAnFld(J) = CurrentDb.QueryDefs(Qry_or_Tbl_Nm).Fields(J).Name
'    Next
'    Exit Function
'End If
'ss.A 1, "Given name is not table or query"
'E: Fnd_AyFld = True: ss.B cSub, cMod, "Qry_or_Tbl_Nm", Qry_or_Tbl_Nm
'End Function
'Function Fnd_AyFn_ByLik(oAyFn$(), pDir$, pLik$, Optional pNoExt As Boolean = False) As Boolean
'Const cSub$ = "Fnd_AyFn_ByLik"
''Aim: Fnd {oAyFn} by {pLik} in {pDir}
'If Not IsDir(pDir) Then ss.A 1: GoTo E
'Dim mFn$, mAyFn$(), N As Byte
'mFn = VBA.Dir(pDir & "*.*")
'While mFn <> ""
'    If mFn Like pLik Then
'        ReDim Preserve mAyFn(N): N = N + 1
'        If pNoExt Then
'            mAyFn(N - 1) = Cut_Ext(mFn)
'        Else
'            mAyFn(N - 1) = mFn
'        End If
'    End If
'    mFn = VBA.Dir
'Wend
'oAyFn = mAyFn
'Exit Function
'R: ss.R
'E: Fnd_AyFn_ByLik = True: ss.B cSub, cMod, ""
'End Function
'Function Fnd_AyFn(oAyFn$(), pDir$, Optional pFspc$ = "*.xls", Optional pNoExt As Boolean = False) As Boolean
''Aim: Fnd {oAyFn} by {pFSpc} in {pDir}
'Const cSub$ = "Fnd_AyFn"
'If Not IsDir(pDir) Then ss.A 1: GoTo E
'Dim mFn$, mAyFn$(), N As Byte, mAyLik$(): mAyLik = Split(pFspc, ",")
'mFn = VBA.Dir(pDir & "*.*")
'While mFn <> ""
'    If IsLikAyLik(mFn, mAyLik) Then
'        ReDim Preserve mAyFn(N): N = N + 1
'        If pNoExt Then
'            mAyFn(N - 1) = Cut_Ext(mFn)
'        Else
'            mAyFn(N - 1) = mFn
'        End If
'    End If
'    mFn = VBA.Dir
'Wend
'oAyFn = mAyFn
'Exit Function
'R: ss.R
'E: Fnd_AyFn = True: ss.B cSub, cMod, "pDir,pFspc,pNoExt"
'End Function
'Function Fnd_An2V_ByFrm(oAn2V() As tNm2V, pFrm As Access.Form, FnStr$) As Boolean
''Aim: Fnd {oAn2V} from {FnStr} in {pFrm} with optional to replace the {.Nm} of {oAn2V} by {pLnNew}
'Const cSub$ = "Fnd_AnV_ByFrm"
'Dim mAn_Frm$(), mAn_Host$(): If Brk_Lm_To2Ay(mAn_Frm, mAn_Host, FnStr) Then ss.A 1: GoTo E
'Dim N%: N = Siz_Ay(mAn_Frm): If N = 0 Then Exit Function
'ReDim oAn2V(N - 1)
'On Error GoTo R
'Dim J%, iCtl As Access.Control
'For J = 0 To N - 1
'    If Fnd_Ctl(iCtl, pFrm, mAn_Frm(J)) Then ss.A 2: GoTo E
'    With oAn2V(J)
'        .Nm = mAn_Host(J)
'        .NewV = iCtl.Value
'        .OldV = iCtl.OldValue
'    End With
'Next
'Exit Function
'R: ss.R
'E: Fnd_An2V_ByFrm = True: ss.B cSub, cMod, "pFrm,FnStr", ToStr_Frm(pFrm), FnStr
'End Function
'Function Fnd_An2V_ByFrm__Tst()
'Const cNmFrm$ = "frmSelBrandEnv"
'If FrmOpn(cNmFrm) Then Stop: GoTo E
'Dim mFrm As Access.Form: Set mFrm = Access.Application.Forms(cNmFrm)
'Dim mAyNm2V() As tNm2V: If Fnd_An2V_ByFrm(mAyNm2V, mFrm, "") Then Stop
'Stop
'Exit Function
'E: Fnd_An2V_ByFrm_Tst = True
'End Function
'Function Fnd_Anm_ByPrj(oAnm$(), pPrj As vbproject _
'    , Optional pLikNmm$ = "*" _
'    , Optional pSrt As Boolean = False _
'    ) As Boolean
'Const cSub$ = "Fnd_Anm_ByPrj"
'Clr_Ays oAnm
'With pPrj
'    Dim mCmp As VBIDE.VBComponent
'    For Each mCmp In .VBComponents
'        If mCmp.Name Like pLikNmm Then Add_AyEle oAnm, mCmp.Name
'    Next
'End With
'If pSrt Then If Srt_Ay(oAnm, oAnm) Then ss.A 1: GoTo E
'Exit Function
'R: ss.R
'E: Fnd_Anm_ByPrj = True: ss.B cSub, cMod, "pPrj,pLikNmm,pSrt", ToStr_Prj(pPrj), pLikNmm, pSrt
'End Function

'Function Fnd_Anm__Tst()
'Dim mAnPrj$(): If Fnd_AnPrj(mAnPrj) Then Stop: GoTo E
'Dim J%
'For J = 0 To Siz_Ay(mAnPrj) - 1
'    Dim mPrj As vbproject: If Fnd_Prj(mPrj, mAnPrj(J)) Then Stop: GoTo E
'    Dim mAnm$(): If Fnd_Anm_ByPrj(mAnm, mPrj) Then Stop: GoTo E
'    Debug.Print mAnPrj(J) & ": " & ToStr_Ays(mAnm)
'Next
'Exit Function
'E: Fnd_Anm_Tst = True
'End Function

'Function Fnd_AnObj_ByPfx(oAnObj$(), pPfx$, Optional pTypObj As Access.AcObjectType = Access.AcObjectType.acQuery) As Boolean
'Const cSub$ = "Fnd_AnObj_ByPfx"
'If Fnd_AnObj_ByPfx_InMdb(oAnObj, "", pPfx, pTypObj) Then ss.A 1: GoTo E
'Exit Function
'R: ss.R
'E: Fnd_AnObj_ByPfx = True: ss.B cSub, cMod, "pPfx,pTypObj", pPfx, ToStr_TypObj(pTypObj)
'End Function
'Function Fnd_AnObj_ByPfx_InMdb(oAnObj$(), pFb$, pPfx$, Optional pTypObj As Access.AcObjectType = Access.AcObjectType.acQuery) As Boolean
'Const cSub$ = "Fnd_AnObj_ByPfx_InMdb"
'Select Case pTypObj
'Case Access.AcObjectType.acQuery
'    If pFb = "" Or pFb = CurrentDb.Name Then
'        If Fnd_Anq_ByPfx(oAnObj, pPfx) Then ss.A 1: GoTo E
'    Else
'        Dim mDb As database: If Opn_Db_R(mDb, pFb) Then ss.A 2: GoTo E
'        If Fnd_Anq_ByPfx(oAnObj, pPfx, mDb) Then ss.A 3: GoTo E
'        Cls_Db mDb
'    End If
'    Exit Function
'End Select
'ss.A 4, "At this moment, only Query Type is supported": GoTo E
'R: ss.R
'E: Fnd_AnObj_ByPfx_InMdb = True: ss.B cSub, cMod, "pFb,pPfx,pTypObj", pFb, pPfx, ToStr_TypObj(pTypObj)
'End Function
'Function Fnd_AnPrj(oAnPrj$() _
'    , Optional pLikNmPrj$ = "*" _
'    , Optional pSrt As Boolean = False _
'    , Optional pAcs As Access.Application = Nothing _
'    ) As Boolean
'Clr_Ays oAnPrj
'Dim mAcs As Access.Application: Set mAcs = Cv_Acs(pAcs)
'Dim iPrj As vbproject
'For Each iPrj In mAcs.Vbe.VBProjects
'    If iPrj.Name Like pLikNmPrj Then Add_AyEle oAnPrj, iPrj.Name
'Next
'End Function

'Function Fnd_AnPrj__Tst()
'Dim mAnPrj$(): If Fnd_AnPrj(mAnPrj) Then Stop: GoTo E
'Debug.Print ToStr_Ays(mAnPrj, , vbLf)
'Shw_DbgWin
'Exit Function
'E: Fnd_AnPrj_Tst = True
'End Function

'Function Fnd_Anq_ByNmQs(oAnq$(), QryNms$ _
'    , Optional pMajBeg As Byte = 0 _
'    , Optional pMajEnd As Byte = 99 _
'    , Optional pDbQry As database) As Boolean
'Dim mDb As database: Set mDb = DbNz(pDbQry)
'Dim L As Byte: L = Len(QryNms) + 1
'Dim mNmQs$: mNmQs = QryNms & "_"
'Clr_Ays oAnq
'
'Dim mMajBeg$: mMajBeg$ = Format(pMajBeg, "00")
'Dim mMajEnd$: mMajEnd$ = Format(pMajEnd, "00") & Chr(255)
'Dim I%
'Dim iQry As QueryDef: For Each iQry In mDb.QueryDefs
'    If Left(iQry.Name, L) <> mNmQs Then GoTo Nxt
'    If iQry.Name < QryNms & "_" & mMajBeg$ Then GoTo Nxt
'    If iQry.Name > QryNms & "_" & mMajEnd$ Then Exit For
'    ReDim Preserve oAnq(I): oAnq(I) = iQry.Name: I = I + 1
'Nxt:
'Next
'End Function
'Function Fnd_Anq_ByNmqs__Tst()
'Const cSub$ = "Fnd_Anq_ByNmqs_Tst"
'Dim mAy$(), mNmQs$
'Dim mResult As Boolean
'Dim mCase As Byte: mCase = 1
'Select Case mCase
'Case 1
'    mNmQs = "qryRfhInqAR"
'    mResult = Fnd_Anq_ByNmQs(mAy$, mNmQs, , 7)
'End Select
'Shw_Dbg cSub, cMod, , "mResult,mNmqs,mAy", mResult, mNmQs, ToStr_Ays(mAy)
'End Function
'Function Fnd_Anq_ByPfx(oAnq$(), pPfx$, Optional pDb As database) As Boolean
'Const cSub$ = "Fnd_Anq_ByPfx"
'Dim mDb As database: Set mDb = DbNz(pDb)
'Dim L%: L% = Len(pPfx)
'Dim I%
'Clr_Ays oAnq
'Dim mPfxXX$: mPfxXX = pPfx & Chr(255)
'Dim iQry As DAO.QueryDef: For Each iQry In mDb.QueryDefs
'    If Left(iQry.Name, L) = pPfx Then
'        ReDim Preserve oAnq$(I)
'        oAnq(I) = iQry.Name
'        I = I + 1
'    End If
'    If iQry.Name > mPfxXX Then Exit For
'Next
'End Function
'Function Fnd_Ant_ByLnk(oAnt_Lnk$(), Optional pDb As database) As Boolean
'Const cSub$ = "Fnd_Ant_ByLnk"
'Dim mDb As database: Set mDb = DbNz(pDb)
'Dim I%
'Clr_Ays oAnt_Lnk
'Dim iTbl As DAO.TableDef: For Each iTbl In mDb.TableDefs
'    If iTbl.Connect <> "" Then
'        ReDim Preserve oAnt_Lnk(I)
'        oAnt_Lnk(I) = iTbl.Name
'        I = I + 1
'    End If
'Next
'End Function

'Function Fnd_Ant_ByLnk__Tst()
'Dim mAnt_Lnk$()
'If Fnd_Ant_ByLnk(mAnt_Lnk) Then Stop
'Debug.Print Join(mAnt_Lnk, vbLf)
'End Function

'Function Fnd_Ant_ByLik(oAnt$(), pLik$, Optional pDb As database, Optional pQ$ = "") As Boolean
'Const cSub$ = "Fnd_Ant_ByLik"
'Dim mDb As database: Set mDb = DbNz(pDb)
'Dim I%
'Clr_Ays oAnt
'Dim iTbl As DAO.TableDef: For Each iTbl In mDb.TableDefs
'    Dim mNmt$: mNmt = iTbl.Name
'    If Left(mNmt, 4) <> "MSYS" Then
'        If iTbl.Name Like pLik$ Then
'            ReDim Preserve oAnt$(I)
'            oAnt(I) = Q_S(iTbl.Name, pQ)
'            I = I + 1
'        End If
'    End If
'Next
'End Function
'Function Fnd_Anq_ByLik(oAnq$(), pLik$, Optional pDb As database, Optional pQ$ = "") As Boolean
'Const cSub$ = "Fnd_Anq_ByLik"
'Dim mDb As database: Set mDb = DbNz(pDb)
'Dim I%
'Clr_Ays oAnq
'Dim iQry As DAO.QueryDef: For Each iQry In mDb.QueryDefs
'    Dim mNmq$: mNmq = iQry.Name
'    If Left(mNmq, 1) <> "~" Then
'        If mNmq Like pLik$ Then
'            ReDim Preserve oAnq$(I)
'            oAnq(I) = Q_S(mNmq, pQ)
'            I = I + 1
'        End If
'    End If
'Next
'End Function
'Function Fnd_An_BySetNm(oAn$(), pAn$(), pSetNm$) As Boolean
''Aim:
'End Function
'Function Fnd_Antq_ByLik(oAntq$(), pLik$, Optional pDb As database, Optional pQ$ = "") As Boolean
'Const cSub$ = "Fnd_Antq_ByLik"
'Dim mDb As database: Set pDb = DbNz(mDb)
'Dim I%
'Clr_Ays oAntq
'Dim mAnt$(), mAnq$()
'If Fnd_Ant_ByLik(mAnt, pLik, mDb, pQ) Then ss.A 1: GoTo E
'If Fnd_Anq_ByLik(mAnq, pLik, mDb, pQ) Then ss.A 2: GoTo E
'If Add_Ay(oAntq, mAnt, mAnq) Then GoTo E
'Exit Function
'E: Fnd_Antq_ByLik = True: ss.B cSub, cMod, "pLik,pDb,pQ", pLik, ToStr_Db(pDb), pQ
'End Function
'Function Fnd_AnWs_BySetWs(oAnWs$(), pWb As Workbook, Optional pSetWs$ = "*") As Boolean
'Const cSub$ = "Fnd_AnWs_BySetWs"
'On Error GoTo R
'Dim mAnLikNmWs$(): mAnLikNmWs = Split(pSetWs, CtComma)
'Dim J%, mAnWs$()
'For J = 0 To Siz_Ay(mAnLikNmWs) - 1
'    Dim mLikNmWs$: If Fnd_AnWs_ByLikNmWs(mAnWs, pWb, Trim(mAnLikNmWs(J))) Then ss.A 1: GoTo E
'    If Add_AyAtEnd(oAnWs, mAnWs) Then ss.A 2: GoTo E
'Next
'Exit Function
'R: ss.R
'E: Fnd_AnWs_BySetWs = True: ss.B cSub, cMod, "pWb,pSetWs", ToStr_Wb(pWb), pSetWs
'End Function
'Function Fnd_AnWs_ByLikNmWs(oAnWs$(), pWb As Workbook, pLikNmWs$) As Boolean
'Const cSub$ = "Fnd_AnWs_ByLikNmWs"
'On Error GoTo R
'If InStr(pLikNmWs, "*") = 0 Then
'    If IsWs(pWb, pLikNmWs) Then
'        ReDim oAnWs(0): oAnWs(0) = pLikNmWs
'        Exit Function
'    End If
'    Dim mA$(): oAnWs = mA
'    Exit Function
'End If
'Dim iWs As Worksheet, mN%: mN = 0
'For Each iWs In pWb.Sheets
'    If iWs.Name Like pLikNmWs Then
'        ReDim Preserve oAnWs(mN)
'        oAnWs(mN) = iWs.Name
'        mN = mN + 1
'    End If
'Next
'Exit Function
'R: ss.R
'E: Fnd_AnWs_ByLikNmWs = True: ss.B cSub, cMod, "pWb,pLikNmWs", ToStr_Wb(pWb), pLikNmWs
'End Function
'Function Fnd_AnWs(oAnWs$(), pFx$, Optional pInclInvisible As Boolean = False) As Boolean
'Const cSub$ = "Fnd_AnWs"
'Dim mWb As Workbook, iWs As Worksheet, J As Byte
'If Opn_Wb(mWb, pFx, True) Then ss.A 1: GoTo E
'If Fnd_AnWs_ByWb(oAnWs, mWb, pInclInvisible) Then ss.A 2: GoTo E
'mWb.Close False
'Exit Function
'R: ss.R
'E: Fnd_AnWs = True: ss.B cSub, cMod, "pFx,pInclInvisible", pFx, pInclInvisible
'End Function
'Function Fnd_AnWs_ByWb(oAnWs$(), pWb As Workbook, Optional pInclInvisible As Boolean = False) As Boolean
'Const cSub$ = "Fnd_AnWs"
'On Error GoTo R
'ReDim oAnWs$(pWb.Sheets.Count - 1)
'Dim J%, iWs As Worksheet: J = 0
'For Each iWs In pWb.Sheets
'    If Not pInclInvisible And Not iWs.Visible Then GoTo Nxt
'    oAnWs(J) = iWs.Name: J = J + 1
'Nxt:
'Next
'Exit Function
'R: ss.R
'E: Fnd_AnWs_ByWb = True: ss.B cSub, cMod, "pWb,pInclinvisble", ToStr_Wb(pWb), pInclInvisible
'End Function
'Public Function Fnd_AnWs_wColr(oAnWs$(), pWb As Workbook) As Boolean
''Aim: Find {oAnws} with color in tab
'If TypeName(pWb) = "Nothing" Then Fnd_AnWs_wColr = True: Exit Function
'Dim iWs As Worksheet, iCnt As Byte
'For Each iWs In pWb.Sheets
'    If iWs.Tab.Color Then
'        ReDim Preserve oAnWs(iCnt)
'        oAnWs(iCnt) = iWs.Name
'        iCnt = iCnt + 1
'    End If
'Next
'If iCnt = 0 Then Fnd_AnWs_wColr = True
'End Function
'Function Fnd_Brand_ById(oBrand$, pBrandId As Byte) As Boolean
'Const cSub$ = "Fnd_Brand_ById"
'Dim mRs As DAO.Recordset: If Opn_Rs(mRs, "Select Brand from mstBrand where BrandId=" & pBrandId) Then ss.A 1: GoTo E
'With mRs
'    If .AbsolutePosition = -1 Then .Close: ss.A 2, "No record such {pBrandId} in mstBrand": GoTo E
'    If Nz(!Brand, "") = "" Then .Close: ss.A 3, "Empty value in [Brand] field is found in mstBrand": GoTo E
'    oBrand = !Brand
'    .Close
'End With
'Exit Function
'R: ss.R
'E: Fnd_Brand_ById = True: ss.B cSub, cMod, "pBrandId", pBrandId
'End Function
''Function Fnd_CdMd(oMod As CodeModule, pMod$) As Boolean
''Const cSub$ = "Fnd_CdMd"
''Dim mNmPrj$, mNmm$: If Brk_Str_Both(mNmPrj, mNmm, pMod, ".") Then ss.A 1: GoTo E
''Dim mVBPrj As VBProject: If Fnd_Prj(mVBPrj, mNmPrj) Then ss.A 2: GoTo E
''Dim mVBCmp As VBComponent: If Fnd_VBCmp(mVBCmp, mVBPrj, mNmm) Then ss.A 3: GoTo E
''Set oMod = mVBCmp.CodeModule
''Exit Function
''R: ss.A 255,Err.Description, eException
''
''E: Fnd_
'':ss.B cSub, cMod, ""
''    CdMd = True
''End Function
''Function Fnd_CdMod(oCdMod As CodeModule, pMod$) As Boolean
''Const cSub$ = "Fnd_CdMod"
''Dim mNmPrj$, mNmm$: If Brk_Str_Both(mNmPrj, mNmm, pMod, ".") Then ss.A 1, "pMod must be xx.xx": GoTo E
''Dim mVBPrj As VBProject: If Fnd_Prj(mVBPrj, mNmPrj) Then ss.A 2: GoTo E
''Dim iCmp As vbide.VBComponent: For Each iCmp In mVBPrj.VBComponents
''    If iCmp.Type = vbext_ct_StdModule Then If iCmp.Name = mNmm Then Set oCdMod = iCmp.CodeModule: Exit Function
''Next
''ss.A 3, "CdMod not found": GoTo E
''Exit Function
''R: ss.A 255,Err.Description, eException
''
''E: Fnd_
'':ss.B cSub, cMod, "pMod", pMod
''    CdMod = True
''End Function
''Function Fnd_CdMod__Tst()
''Const cSub$ = "Fnd_CdMod_Tst"
''Dim mCdMod As CodeModule, mMod$
''Dim mRslt As Boolean, mCase As Byte
''mCase = 1
''Select Case mCase
''Case 1
''    mMod$ = "Fnd"
''End Select
''mRslt = Fnd_CdMod(mCdMod, mMod$)
''Shw_Dbg cSub, cMod, , "mRslt,mMod$", mRslt, mMod$
''End Function
'Function Fnd_Cno_XInRow%(pWs As Worksheet, pRno&, Optional pLookFor$ = "X", Optional pCnoFm As Byte = 1, Optional pCnoTo As Byte = 255)
'Dim iCno%, mStp%
'mStp = IIf(pCnoTo >= pCnoFm, 1, -1)
'For iCno = pCnoFm To pCnoTo Step mStp
'    If pWs.Cells(pRno, iCno).Value = pLookFor$ Then Fnd_Cno_XInRow = iCno: Exit Function
'Next
'End Function
'Function Fnd_Cno_EmptyCell_InRow(pWs As Worksheet _
'    , Optional pRno& = 1 _
'    , Optional pCnoFm% = 1 _
'    , Optional pCnoTo% = 256 _
'    ) As Byte
'Dim iCno%, mStp%
'mStp = IIf(pCnoTo >= pCnoFm, 1, -1)
'For iCno = pCnoFm To pCnoTo Step mStp
'    If IsEmpty(pWs.Cells(pRno, iCno).Value) Then Fnd_Cno_EmptyCell_InRow = iCno: Exit Function
'Next
'Fnd_Cno_EmptyCell_InRow = 0
'End Function
'Function Fnd_Ctl(oCtl As Access.Control, pFrm As Access.Form, pNmCtl$) As Boolean
'Const cSub$ = "Fnd_Ctl"
'On Error GoTo R
'Set oCtl = pFrm.Controls(pNmCtl)
'Exit Function
'R: ss.R
'E: Fnd_Ctl = True: ss.B cSub, cMod, "pFrm,pNmCtl", ToStr_Frm(pFrm), pNmCtl
'End Function
'Function Fnd_Env_ById(oEnv$, pEnvId As Byte) As Boolean
'Const cSub$ = "Fnd_Env_ById"
'Dim mRs As DAO.Recordset: If Opn_Rs(mRs, "Select Env from mstEnv where EnvId=" & pEnvId) Then ss.A 1: GoTo E
'With mRs
'    If .AbsolutePosition = -1 Then .Close: ss.A 2, "No such record {pEnvId} in mstEnv": GoTo E
'    If Nz(!Env, "") = "" Then .Close: ss.A 3, "Empty value in [Env] field is found in mstEnv": GoTo E
'    oEnv = !Env
'    .Close
'End With
'Exit Function
'R: ss.R
'E: Fnd_Env_ById = True: ss.B cSub, cMod, "pEnvId", pEnvId
'End Function
'Function Fnd_FctNam_ByNmq$(QryNm$)
''Assume: Return QQQQ_n_xxxx_RunCode from {QryNm} of format QQQQ_nn_n_xxxx_RunCode
'Dim mP1%: mP1 = InStr(QryNm, "_")
'Dim mP2%: mP2 = InStr(mP1 + 1, QryNm, "_")
'If mP1 = 0 Or mP2 = 0 Then Stop
'If mP1 > mP2 Then Stop
'If Right(QryNm, 8) = "_RUNCODE" Then
'    Fnd_FctNam_ByNmq = Left(QryNm, mP1 - 1) & mID$(QryNm, mP2)
'ElseIf Right(QryNm, 4) = "_Run" Then
'    Dim mP3%: mP3 = InStr(mP2 + 1, QryNm, "_")
'    If mP3 = 0 Then Stop
'    Fnd_FctNam_ByNmq = Left(QryNm, mP1 - 1) & mID$(QryNm, mP3)
'Else
'    Stop
'End If
'End Function

'Function Fnd_FctNam_ByNmq__Tst()
'Debug.Print Fnd_FctNam_ByNmq("qryABC_01_1_lsdf_RunCode")
'End Function

'Function Fnd_Ffn(oFfn$, Optional pDir$ = "C:\", Optional pFspc$ = "*.*", Optional pNmFSpc$ = "Any File", Optional pTit$ = "Select a file") As Boolean
'Const cSub$ = "Fnd_Ffn"
'With Application.FileDialog(msoFileDialogFilePicker)
'    .InitialFileName = pDir
'    .AllowMultiSelect = False
'    .Title = pTit
'    .Filters.Add pNmFSpc, pFspc
'    .Show
'    If .SelectedItems.Count = 1 Then oFfn = .SelectedItems(1): Exit Function
'End With
'E: Fnd_Ffn = True
'End Function

'Function Fnd_Ffn__Tst()
'Const cSub$ = "Fnd_Ffn_Tst"
'Dim mFfn$: If Fnd_Ffn(mFfn) Then Stop: GoTo E
'Shw_Dbg cSub, cMod, "mFfn", mFfn
'Exit Function
'E: Fnd_Ffn_Tst = True
'End Function

'Function Fnd_Fb_FmCnnStr(oFb$, pCnnStr$) As Boolean
'Const cSub$ = "Fnd_Fb_FmCnnStr"
'Const cDtaSrc$ = "Data Source="
''Provider=Microsoft.Jet.OLED4.0;User ID=Admin;Data Source=M:\07 ARCollection\ARCollection\WorkingDir\PgmObj\Template_ARInq.mdb;Mode=ReadWrite;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False
'Dim mP1%: mP1 = InStr(pCnnStr, cDtaSrc)
'Dim mP2%: mP2 = InStr(mP1, pCnnStr, ";")
'If mP1 = 0 Or mP2 = 0 Or mP1 > mP2 Then ss.A 1, "Cannot find Data Source= or ; in given connection string": GoTo E
'Dim L As Byte: L = Len(cDtaSrc)
'oFb = mID(pCnnStr, mP1 + L, mP2 - mP1 - L)
'Exit Function
'E: Fnd_Fb_FmCnnStr = True: ss.B cSub, cMod, "pCnnStr", pCnnStr
'End Function
'Function Fnd_Fb_FmNmt_Lnk(oFb$, pNmt_Lnk$) As Boolean
'Const cSub$ = "Fnd_Fb_FmNmt_Lnk"
'Const cDb$ = "DATABASE="
''DATABASE=D:\SPLHalfWayHouse\MPSDetail\VerNew@2007_01_04\WorkingDir\PgmObj\MPS_RfhFc.Mdb;TABLE=tblFcPrm
'On Error GoTo R
'Dim L As Byte: L = Len(cDb)
'Dim mCnnStr$: mCnnStr = CurrentDb.TableDefs(pNmt_Lnk).Connect
'If Left(mCnnStr, 1) = ";" Then mCnnStr = mID(mCnnStr, 2)
'If Left(mCnnStr, L) <> cDb Then ss.A 1, "pNmt_Lnk should have connect string started with " & cDb, , "pNmt_Lnk,CnnStr", pNmt_Lnk, mCnnStr: GoTo E
'Dim mP%: mP = InStr(mCnnStr, ";")
'If mP <= 0 Then oFb = mID(mCnnStr, L + 1): Exit Function
'oFb = mID(mCnnStr, L + 1, mP - L)
'Exit Function
'R: ss.R
'E: Fnd_Fb_FmNmt_Lnk = True: ss.B cSub, cMod, "pNmt_Lnk,mCnnStr", pNmt_Lnk, mCnnStr
'End Function

'Function Fnd_Fb_FmNmt_Lnk__Tst()
'Dim mFb$: If Fnd_Fb_FmNmt_Lnk(mFb, "tblFcPrm") Then Stop
'Debug.Print mFb
'End Function

'Function Fnd_FirstDateOfWk(pYr As Byte, pWk As Byte) As Date
'Dim mFirstDateOfWk1 As Date
'Select Case pYr
'    Case 5:     mFirstDateOfWk1 = #1/2/2005#
'    Case 6:     mFirstDateOfWk1 = #1/1/2006#
'    Case 7:     mFirstDateOfWk1 = #1/7/2007#
'    Case Else
'        Stop
'End Select
'Fnd_FirstDateOfWk = mFirstDateOfWk1 + (pWk - 1) * 7
'End Function
'Function Fnd_FldVal_ByFld(oVal, pFld As DAO.Field) As Boolean
'Const cSub$ = "Fnd_FldVal_ByFld"
'On Error GoTo R
'oVal = pFld.Value
'Exit Function
'R: ss.R
'E: Fnd_FldVal_ByFld = True: ss.B cSub, cMod, "pFld", ToStr_Fld(pFld)
'End Function
'Function Fnd_FldVal(oVal, pRs As DAO.Recordset, pNmFldRet$) As Boolean
'Const cSub$ = "Fnd_FldVal"
'On Error GoTo R
'oVal = pRs.Fields(pNmFldRet).Value
'Exit Function
'R: ss.R
'E: Fnd_FldVal = True: ss.B cSub, cMod, "pNmFldRet,pRs", pNmFldRet, ToStr_Rs(pRs)
'End Function

'Function Fnd_FldVal__Tst()
'Dim mRs As DAO.Recordset: Set mRs = CurrentDb.OpenRecordset("Select * from mstBrand")
'Dim mV: If Fnd_FldVal(mV, mRs, "aaa") Then Stop
'mRs.Close
'End Function

'Function Fnd_AyRgeRno(oAyRgeRno() As tRgeRno, Rg As Range) As Boolean
''Aim: find {oAyRgeRno} by each 'block'.  One 'block' one element in {oAyRgeRno}.
''     a 'block' is pCol for of RnoFm & RnoTo having same value.
''     Rg can be single cell, which means from this cell downward until empty cell
''     or a range of vertical cells to find the tRgeRno
'Const cSub$ = "Fnd_AyRgeRno"
'On Error GoTo R
''-- Find mNRow&
'Dim mNRow&
'If Rg.Count = 1 Then
'    If Rg.End(xlDown).Row = 65536 Then
'        ReDim oAyRgeRno(0)
'        oAyRgeRno(0).Fm = Rg.Row
'        oAyRgeRno(0).To = Rg.Row
'        Exit Function
'    End If
'    mNRow = Rg.End(xlDown).Row - Rg.Row + 1
'Else
'    mNRow = Rg.SpecialCells(xlCellTypeLastCell).Row - Rg.Row + 1
'End If
'
'Dim iRno&, mRnoFm&: mRnoFm = Rg.Row
'Dim mN%: mN = 0
'Dim mWs As Worksheet: Set mWs = Rg.Parent
'Dim mCno As Byte: mCno = Rg.Column
'Dim mV, mVLas
'mVLas = Rg.Cells(1, 1)
'For iRno = Rg.Row To Rg.Row + mNRow - 1
'    If mVLas <> mWs.Cells(iRno, mCno).Value Then
'        mVLas = mWs.Cells(iRno, mCno).Value
'        ReDim Preserve oAyRgeRno(mN)
'        With oAyRgeRno(mN)
'            .Fm = mRnoFm
'            .To = iRno - 1
'        End With
'        mN = mN + 1
'        mRnoFm = iRno
'    End If
'Next
'If iRno > mRnoFm Then
'    ReDim Preserve oAyRgeRno(mN)
'    With oAyRgeRno(mN)
'        .Fm = mRnoFm
'        .To = iRno - 1
'    End With
'End If
'Exit Function
'R: ss.R
'E: Fnd_AyRgeRno = True: ss.B cSub, cMod, "Rg", ToStr_Rge(Rg)
'End Function
'Function Fnd_FmtDefSq(oFmtDefSq As tSq, pQt As QueryTable) As Boolean
'Const cSub$ = "Fnd_FmtDefSq"
'Const cTbl$ = "<Tbl>"
'
'Clr_Sq oFmtDefSq
'
'oFmtDefSq.c1 = pQt.Destination.Column
'Dim mWs As Worksheet: Set mWs = pQt.Parent
'
'Dim mRgeRnoSearch As tRgeRno
'With mRgeRnoSearch
'    .To = pQt.Destination.Row - 1
'    .Fm = .To - 30: If .Fm <= 0 Then .Fm = 1
'    Dim iRno&: For iRno = .Fm To .To
'        If mWs.Cells(iRno, oFmtDefSq.c1).Value = cTbl Then
'            oFmtDefSq.r1 = iRno
'            Dim jRno&: For jRno = iRno + 1 To .To
'                If mWs.Cells(jRno, oFmtDefSq.c1).Value = cTbl Then
'                    oFmtDefSq.r2 = jRno
'                    Exit For
'                End If
'            Next
'            Exit For
'        End If
'    Next
'End With
'
'With oFmtDefSq
'    If .r1 = 0 Or .r2 = 0 Then ss.A 1, "No <Tbl> defintion": GoTo E
'End With
'
'Dim iCno As Byte: For iCno = oFmtDefSq.c1 + 1 To 255
'    If mWs.Cells(oFmtDefSq.r1, iCno).Value = cTbl Then
'        oFmtDefSq.c2 = iCno
'        Exit For
'    End If
'Next
'With oFmtDefSq
'    If .c2 = 0 Then ss.A 2, "No <Tbl> defintion": GoTo E
'End With
'Exit Function
'E: Fnd_FmtDefSq = True: ss.B cSub, cMod, "pQt", ToStr_Qt(pQt)
'End Function
'Public Function Fnd_FreezedCell(oRno&, oCno As Byte, pWs As Worksheet) As Boolean
''Aim: Find the {oFreezedCellAdr} of {pWs}
'Fnd_FreezedCell = True
'pWs.Activate
'Dim mWb As Workbook: Set mWb = pWs.Parent
'Dim mWin As Window: Set mWin = pWs.Application.ActiveWindow
'If mWin.Panes.Count <> 4 Then MsgBox "Given worksheet [" & pWs.Name & "] does not have 4 panes to find the Freezed Cell", vbCritical:  Exit Function
'Dim mA$: mA = mWin.Panes(1).VisibleRange.Address
'Dim mP%: mP = InStr(mA, ":")
'mA = mID(mA, mP + 1)
'Dim mRge As Range: Set mRge = pWs.Range(mA)
'oRno = mRge.Row
'oCno = mRge.Column
'Fnd_FreezedCell = False
'End Function
'Private Function Fnd_FreezedCell__Tst()
''Debug.Print Application.Workbooks.Count
''Debug.Print Application.Workbooks(1).FullName
''Dim mWs As Worksheet: Set mWs = Application.Workbooks(1).Sheets("Input - HKDP")
''Dim mRno&, mCno As Byte: If Fnd_FreezedCell(mRno, mCno, mWs) Then Stop
''Debug.Print mRno & CtComma & mCno
''Dim mSqLeft As cSq, mSqTop As cSq
'End Function
'Function Fnd_NxtBkFfnn(pFfnn$, oNxtBkFfnn$, oNxtBkNo As Byte) As Boolean
'Dim mNmBk$: mNmBk = Right(pFfnn, 10)
'If Left(mNmBk, 8) = " backup(" And Right(mNmBk, 1) = ")" Then
'    oNxtBkNo = Val(mID(mNmBk, 9, 1)) + 1
'    oNxtBkFfnn = Left(pFfnn, Len(pFfnn) - 2) & oNxtBkNo & ")"
'    Exit Function
'End If
'oNxtBkNo = 1
'oNxtBkFfnn = pFfnn & " backup(1)"
'End Function
'Function Fnd_Id_ByBrand(oBrandId As Byte, pBrand$) As Boolean
'Const cSub$ = "Fnd_Id_ByBrand"
'Dim mRs As DAO.Recordset: If Opn_Rs(mRs, "Select BrandId from mstBrand where Brand='" & pBrand & CtSngQ) Then ss.A 1: GoTo E
'With mRs
'    If .AbsolutePosition = -1 Then .Close: ss.A 2, "No such record {pBrand} in mstBrand": GoTo E
'    If Nz(!BrandId, 0) = 0 Then .Close: ss.A 3, "0 in [Brand] field is found in mstBrand": GoTo E
'    oBrandId = !BrandId
'    .Close
'End With
'Exit Function
'R: ss.R
'E: Fnd_Id_ByBrand = True: ss.B cSub, cMod, "pBrand", pBrand
'End Function
'Function Fnd_Id_ByEnv(oEnvId As Byte, pEnv$) As Boolean
'Const cSub$ = "Fnd_Id_ByEnv"
'Dim mRs As DAO.Recordset: If Opn_Rs(mRs, "Select EnvId from mstEnv where Env='" & pEnv & CtSngQ) Then ss.A 1: GoTo E
'With mRs
'    If .AbsolutePosition = -1 Then .Close: ss.A 2, "No record {pEnv} in mstEnv": GoTo E
'    If Nz(!EnvId, 0) = 0 Then .Close: ss.A 3, "0 in [Env] field is found in mstEnv": GoTo E
'    oEnvId = !EnvId
'    .Close
'End With
'Exit Function
'R: ss.R
'E: Fnd_Id_ByEnv = True: ss.B cSub, cMod, "pEnv", pEnv
'End Function
'Function Fnd_LasDteOfLasWW(pDte As Date) As Date
'Dim mWeekday As Byte: mWeekday = Weekday(pDte, vbSunday) ' Sunday count as first day of a week & week day of Sunday is 1 & Saturday (last date of a week) is 7
'Fnd_LasDteOfLasWW = pDte - mWeekday
'End Function
'Function Fnd_Layout(oLnFld$, pWs As Worksheet) As Boolean
'Const cSub$ = "Fnd_Layout"
'On Error GoTo R
'Dim mV: mV = pWs.Cells(1, 1).Value
'If IsEmpty(mV) Then ss.A 1, "A1 cell is empty", "pWs": GoTo E
'oLnFld = mV
'Dim J%: For J = 2 To 255
'    mV = pWs.Cells(1, J).Value
'    If IsEmpty(mV) Then Exit Function
'    oLnFld = oLnFld & CtComma & mV
'Next
'Exit Function
'R: ss.R
'E: Fnd_Layout = True: ss.B cSub, cMod, "pWs", ToStr_Ws(pWs)
'End Function
'Function Fnd_Lbl(oLbl As Access.Label, pCtl As Access.Control) As Boolean
'Dim mFrm As Access.Form: Set mFrm = pCtl.Parent
'Dim mNm$
'mNm = pCtl.Name & "_Lbl": If Fnd_Lbl_ByNm(oLbl, mFrm, mNm) Then GoTo E
'Exit Function
'E: Fnd_Lbl = True
'End Function
'Function Fnd_Lbl_ByNm(oLbl As Access.Label, pFrm As Access.Form, pNm$) As Boolean
'Const cSub$ = "Fnd_Lbl_ByNm"
'On Error GoTo R
'Set oLbl = pFrm.Controls(pNm)
'Exit Function
'R: Fnd_Lbl_ByNm = True
'End Function
'Function Fnd_Lv_FmDistSql(oLv$, pDistSql$, Optional pDb As database) As Boolean
''Aim: Fnd {oLv} by joinning the first field value of {pDistSql} in {pDb}
'Const cSub$ = "Fnd_Lv_FmDistSql"
'Dim mDb As database: Set mDb = DbNz(pDb)
'With mDb.OpenRecordset(pDistSql)
'    oLv = ""
'    While Not .EOF
'        oLv = Add_Str(oLv, Nz(.Fields(0).Value, "#Null#"), CtComma)
'        .MoveNext
'    Wend
'    .Close
'End With
'Exit Function
'R: ss.R
'E: Fnd_Lv_FmDistSql = True: ss.B cSub, cMod, "pDistSql, mDb", pDistSql, ToStr_Db(mDb)
'End Function
'Function Fnd_Lv_FmIdxTbl(oLv$, IdxNmTbl$, pNmFldRet$, Optional pLn$ = "", Optional pAv = Nothing) As Boolean
''Aim: Fnd {oLv} of {pNmFldRet} in {IdxNmTbl} with filter of list of field of {pLn} with list of value in {pAv}
'Const cSub$ = "Fnd_Lv_FmIdxTbl"
'Dim mWhere$: If Bld_Where(mWhere, pLn, pAv) Then ss.A 1: GoTo E
'Dim mSql$: mSql = Fmt_Str("Select distinct {0} from {1}{2}", pNmFldRet, IdxNmTbl, mWhere)
'If Fnd_Lv_FmDistSql(oLv, mSql) Then ss.A 1: GoTo E
'Exit Function
'R: ss.R
'E: Fnd_Lv_FmIdxTbl = True: ss.B cSub, cMod, "IdxNmTbl,pNmFldRet,pLn,pAv", IdxNmTbl, pNmFldRet, pLn, ToStr_Vayv(pAv)
'End Function
'Function Fnd_LoAyV_FmRs(pRs As DAO.Recordset, FnStr$, oAyV0, Optional oAyV1, Optional oAyV2, Optional oAyV3, Optional oAyV4, Optional oAyV5) As Boolean
'Const cSub$ = "Fnd_LoAyV_FmRs"
'If VarType(oAyV0) And vbArray = 0 Then ss.A 1, "oAyV0 must be an array": GoTo E
'If Not IsMissing(oAyV1) Then If VarType(oAyV1) And vbArray = 0 Then ss.A 2, "oAyV1 must be an array": GoTo E
'If Not IsMissing(oAyV2) Then If VarType(oAyV2) And vbArray = 0 Then ss.A 3, "oAyV2 must be an array": GoTo E
'If Not IsMissing(oAyV3) Then If VarType(oAyV3) And vbArray = 0 Then ss.A 4, "oAyV3 must be an array": GoTo E
'If Not IsMissing(oAyV4) Then If VarType(oAyV4) And vbArray = 0 Then ss.A 5, "oAyV4 must be an array": GoTo E
'If Not IsMissing(oAyV5) Then If VarType(oAyV5) And vbArray = 0 Then ss.A 6, "oAyV5 must be an array": GoTo E
'Dim mAnFld$():  mAnFld = Split(FnStr, CtComma)
'Dim mNFld%:     mNFld = Siz_Ay(mAnFld): If mNFld <= 0 Or mNFld > 6 Then ss.A 7, "FnStr is invalid (at most 6 elements)": GoTo E
'Dim mNRec%:     If Fnd_RecCnt_ByRs(mNRec, pRs) Then ss.A 1: GoTo E
'If mNRec% = 0 Then
'    Dim mAy()
'    oAyV0 = mAy
'    oAyV1 = mAy
'    oAyV2 = mAy
'    oAyV3 = mAy
'    oAyV4 = mAy
'    oAyV5 = mAy
'    Exit Function
'End If
'If Chk_Struct_Rs(pRs, FnStr) Then ss.A 1: GoTo E
'ReDim oAyV0(mNRec - 1), oAyV1(mNRec - 1), oAyV2(mNRec - 1), oAyV3(mNRec - 1), oAyV4(mNRec - 1), oAyV5(mNRec - 1)
'
'With pRs
'    .MoveFirst
'    Dim iRec%: iRec = 0
'    While Not .EOF
'        Dim J%
'        For J = 0 To mNFld - 1
'            Select Case J
'            Case 0: oAyV0(iRec) = .Fields(mAnFld(J)).Value
'            Case 1: oAyV1(iRec) = .Fields(mAnFld(J)).Value
'            Case 2: oAyV2(iRec) = .Fields(mAnFld(J)).Value
'            Case 3: oAyV3(iRec) = .Fields(mAnFld(J)).Value
'            Case 4: oAyV4(iRec) = .Fields(mAnFld(J)).Value
'            Case 5: oAyV5(iRec) = .Fields(mAnFld(J)).Value
'            End Select
'        Next
'        .MoveNext
'        iRec = iRec + 1
'    Wend
'End With
'Exit Function
'R: ss.R
'E: Fnd_LoAyV_FmRs = True: ss.B cSub, cMod, "FnStr,mNRec,mNFld,Rs", FnStr, mNRec, mNFld, ToStr_Rs(pRs)
'End Function
'Function Fnd_LoAyV_FmRs__Tst()
'Dim mRs As DAO.Recordset: Set mRs = CurrentDb.OpenRecordset("Select * from tblOdbcSql")
'Dim mLn$:                 mLn = "NmQs,NmDl" ' ,Sql,Sql_LclMd"
'Dim AnQs(), AnDl(), AySql(), AySql_LclMd()
'Fnd_LoAyV_FmRs_Tst = Fnd_LoAyV_FmRs(mRs, mLn, AnQs, AnDl, AySql, AySql_LclMd)
'Shw_Dbg "GetLoAyV_FmRs_Tst", cMod, , "AnQs,AnDl", ToStr_AyV(AnQs), ToStr_AyV(AnDl)
'End Function
'Function Fnd_LoAyV_FmSql(Sql$, pLn$, oAyV0, Optional oAyV1 As Variant, Optional oAyV2, Optional oAyV3, Optional oAyV4, Optional oAyV5) As Boolean
'Fnd_LoAyV_FmSql = Fnd_LoAyV_FmSql_InDb(CurrentDb, Sql, pLn, oAyV0, oAyV1, oAyV2, oAyV3, oAyV4, oAyV5)
'End Function
'Function Fnd_LoAyV_FmSql_InDb(pDb As database, Sql$, pLn$, oAyV0, Optional oAyV1 As Variant, Optional oAyV2, Optional oAyV3, Optional oAyV4, Optional oAyV5) As Boolean
'Const cSub$ = "Fnd_LoAyV_FmRs"
'Dim mDb As database: Set mDb = DbNz(pDb)
'Dim mRs As DAO.Recordset: If Fnd_Rs_BySql(mRs, Sql, mDb) Then ss.A 1: GoTo E
'If Fnd_LoAyV_FmRs(mRs, pLn, oAyV0, oAyV1, oAyV2, oAyV3, oAyV4, oAyV5) Then ss.A 2: GoTo E
'mRs.Close
'Exit Function
'R: ss.R
'E: Fnd_LoAyV_FmSql_InDb = True: ss.B cSub, cMod, "pDb,Sql,pLn", ToStr_Db(pDb), Sql, pLn
'End Function
'Function Fnd_LnFldVal(pRs As DAO.Recordset, FnStr$ _
'    , Optional oV0 _
'    , Optional oV1 _
'    , Optional oV2 _
'    , Optional oV3 _
'    , Optional oV4 _
'    , Optional oV5 _
'    , Optional oV6 _
'    , Optional oV7 _
'    , Optional oV8 _
'    , Optional oV9 _
'    ) As Boolean
'Const cSub$ = "Fnd_LnFldVal"
'Dim mAnFld$(): If Brk_Ln2Ay(mAnFld, FnStr) Then ss.A 1: GoTo E
'Dim N%: N = Siz_Ay(mAnFld): If N > 10 Then ss.A 1, "No more than 10 fields can be return": GoTo E
'On Error GoTo R
'Dim J%: For J = 0 To N - 1
'    Select Case J
'    Case 0: oV0 = pRs.Fields(mAnFld(J)).Value
'    Case 1: oV1 = pRs.Fields(mAnFld(J)).Value
'    Case 2: oV2 = pRs.Fields(mAnFld(J)).Value
'    Case 3: oV3 = pRs.Fields(mAnFld(J)).Value
'    Case 4: oV4 = pRs.Fields(mAnFld(J)).Value
'    Case 5: oV5 = pRs.Fields(mAnFld(J)).Value
'    Case 6: oV6 = pRs.Fields(mAnFld(J)).Value
'    Case 7: oV7 = pRs.Fields(mAnFld(J)).Value
'    Case 8: oV8 = pRs.Fields(mAnFld(J)).Value
'    Case 9: oV9 = pRs.Fields(mAnFld(J)).Value
'    End Select
'Next
'Exit Function
'R: ss.R
'E: Fnd_LnFldVal = True: ss.B cSub, cMod, "pRs,FnStr,N,J", ToStr_Rs(pRs), FnStr, N, J
'End Function
'Function Fnd_MaxDir$(pDir$)
''Aim: within the all dir in {pDir}, return the dir with Max name
'Dim mMaxDir$
'Dim mDir$: mDir = VBA.Dir(pDir, vbDirectory)
'While mDir <> ""
'    If mDir > mMaxDir Then mMaxDir = mDir
'    mDir = VBA.Dir
'Wend
'Fnd_MaxDir = mMaxDir
'End Function
'Function Fnd_MaxFfn$(pDir$, pFfnSpec$)
''Aim: within the all files of {pFfnSpec} in {pDir}, return the file with Max name
'Dim mMaxFfn$
'Dim mFfn$: mFfn = VBA.Dir(pDir & pFfnSpec)
'While mFfn <> ""
'    If mFfn > mMaxFfn Then mMaxFfn = mFfn
'    mFfn = VBA.Dir
'Wend
'Fnd_MaxFfn = mMaxFfn
'End Function
'Function Fnd_MaxVal(oMaxVal, pNmt$, pNmFldMax$, Optional pLExpr$ = "", Optional pDb As database) As Boolean
'Const cSub$ = "Fnd_MaxVal"
'Dim mDb As database: Set mDb = DbNz(pDb)
'Dim mWhere$: If pLExpr <> "" Then mWhere = " Where " & pLExpr
'On Error GoTo R
'With mDb.OpenRecordset(Fmt_Str("select Max({0}) from {1}{2}", pNmFldMax, Q_SqBkt(pNmt), mWhere))
'    oMaxVal = Nz(.Fields(0).Value, 0)
'    .Close
'End With
'Exit Function
'R: ss.R
'E: Fnd_MaxVal = True: ss.B cSub, cMod, "pNmt,pLExpr,pDb", pNmt, pLExpr, ToStr_Db(pDb)
'End Function
'Public Function Fnd_Nm(oNm As Excel.Name, pWs As Worksheet, pNm$) As Boolean
'If Not Fnd_Nm_InWs(oNm, pWs, pNm) Then Exit Function
'Fnd_Nm = Fnd_Nm_InWb(oNm, pWs.Parent, pNm)
'End Function
'Public Function Fnd_Nm_InWb(oNm As Excel.Name, pWb As Workbook, pNm$) As Boolean
'Const cSub$ = "Fnd_Nm_InWb"
'On Error GoTo R
'Set oNm = pWb.Names(pNm)
'Exit Function
'R: ss.R
'E: Fnd_Nm_InWb = True: ss.B cSub, cMod, "pWb,pNm", ToStr_Wb(pWb), pNm
'End Function
'Public Function Fnd_Nm_InWs(oNm As Excel.Name, pWs As Worksheet, pNm$) As Boolean
'Const cSub$ = "Fnd_Nm_InWs"
'On Error GoTo R
'Set oNm = pWs.Names(pNm)
'Exit Function
'R: ss.R
'E: Fnd_Nm_InWs = True: ss.B cSub, cMod, "pWs,pNm", ToStr_Ws(pWs), pNm
'End Function
'Function Fnd_NmQs$(QryNm$)
''Aim: If the given {QryNm} in format of XXXX_NN_N_xxxx, return XXXX else return ""
'' Postition of first 3 "_"
'Dim mP1 As Byte: mP1 = InStr(QryNm, "_"):         If mP1 <= 1 Then Exit Function
'Dim mP2 As Byte: mP2 = InStr(mP1 + 1, QryNm, "_"): If mP2 <= 0 Then Exit Function
'If mP2 - mP1 <> 3 Then Exit Function
''
'Dim mNN$: mNN = mID$(QryNm, mP1 + 1, 2)
'Dim mA$:  mA = Left(mNN, 1)
'If "0" > mA Or mA > "9" Then Exit Function
'mA = mID$(mNN, 2, 1)
'If "0" > mA Or mA > "9" Then Exit Function
'
'Dim mN$: mN = mID$(QryNm, mP2 + 1, 1)
'mA = mN
'If "0" > mA Or mA > "9" Then Exit Function
'Fnd_NmQs = Left(QryNm, mP1 - 1)
'End Function
'Function Fnd_Nmqs__Tst()
'Debug.Print Fnd_NmQs("qryXCmp")
'End Function
'Function Fnd_Nmt_FmQt(oNmt$, pQt As QueryTable) As Boolean
'Const cSub$ = "Fnd_Nmt_FmQt"
'Select Case pQt.CtCommandType
'Case XlCmdType.xlCmdTable: oNmt = pQt.CtCommandText
'Case XlCmdType.xlCmdSql: If Fnd_Nmt_FmSql(oNmt, pQt.CtCommandText) Then ss.A 1: GoTo E
'Case Else
'    ss.A 1, "Unexpected CmdTyp in given Qt": GoTo E
'End Select
'Exit Function
'R: ss.R
'E: Fnd_Nmt_FmQt = True: ss.B cSub, cMod, "pQt", ToStr_Qt(pQt)
'End Function
'Function Fnd_Nmt_FmSql(oNmt$, Sql$) As Boolean
''Aim Find {oNmt} from {Sql} by looking up the token after "From"
'Const cSub$ = "Fnd_Nmt_FmSql"
'Sql = Replace(Replace(Sql, vbLf, " "), vbCr, " ")
'Dim mAy$(): mAy = Split(Sql, " ")
'Dim J%: For J = 0 To UBound(mAy)
'    If mAy(J) = "From" Then
'        Dim I As Byte: For I = 1 To UBound(mAy)
'            If mAy(J + I) <> "" Then oNmt = Trim(mAy(J + I)): Exit Function
'        Next
'        ss.A 1, "No non-empty element in mAy()", , "mAy", ToStr_Ays(mAy, "[]"): GoTo E
'    End If
'Next
'ss.A 1, "No From in Sql": GoTo E
'Exit Function
'E: Fnd_Nmt_FmSql = True: ss.B cSub, cMod, "Sql", Sql
'End Function
'Function Fnd_Nmt_FmSql__Tst()
'Dim mNmt$: If Fnd_Nmt_FmSql(mNmt, "lkdf lsdkj from    ksd  sdlk") Then Stop
'Debug.Print mNmt
'Stop
'End Function
'Function Fnd_PrcBody_ByMd(oStr$, pMd As CodeModule, pNmPrc$ _
'    , Optional pBodyOnly As Boolean = False _
'    ) As Boolean
'Const cSub$ = "Fnd_PrcBody_ByMd"
'Dim mAnPrc$(): If Fnd_AnPrc_ByMd(mAnPrc, pMd, pNmPrc, , True, pBodyOnly) Then ss.A 1: GoTo E
'If Siz_Ay(mAnPrc) = 0 Then ss.A 2, "pNmPrc is not found": GoTo E
'On Error GoTo R
'Dim iNmPrc$, iPrcLinBeg$, iPrcLinEnd$, iPrcNLin$
'If Brk_Str_To3Seg(iNmPrc, iPrcLinBeg, iPrcLinEnd, mAnPrc(0)) Then ss.A 3: GoTo E
'Dim mNLin&
'Dim mLinBeg&: mLinBeg = iPrcLinBeg
'Dim mLinEnd&: mLinEnd = iPrcLinEnd
'mNLin = mLinEnd - mLinBeg + 1
'oStr = pMd.Lines(mLinBeg, mNLin)
'Exit Function
'R: ss.R
'E: Fnd_PrcBody_ByMd = True: ss.C cSub, cMod, "pMd,pNmPrc,pBodyOnly", ToStr_Md(pMd), pNmPrc$, pBodyOnly
'End Function
'Function Fnd_PrcBody(oStr$, pMod$, pNmPrc$ _
'    , Optional pAcs As Access.Application = Nothing _
'    , Optional pBodyOnly As Boolean = False _
'    ) As Boolean
'Const cSub$ = "Fnd_PrcBody"
'On Error GoTo R
'Dim mNmPrj$, mNmm$: If Brk_Str_Both(mNmPrj, mNmm, pMod, ".") Then ss.A 1, "pMod must have a '.'": GoTo E
'Dim mPrj As vbproject: If Fnd_Prj(mPrj, mNmPrj, pAcs) Then ss.A 2: GoTo E
'Dim mMd As CodeModule: If Fnd_Md(mMd, mPrj, mNmm) Then ss.A 3: GoTo E
'If Fnd_PrcBody_ByMd(oStr, mMd, pNmPrc, pBodyOnly) Then ss.A 4: GoTo E
'Exit Function
'R: ss.R
'E: Fnd_PrcBody = True: ss.C cSub, cMod, "pMod,pNmPrc,pAcs,pBodyOnly", pMod, pNmPrc$, ToStr_Acs(pAcs), pBodyOnly
'End Function
''--------------------

'Function Fnd_PrcBody__Tst()
'Const cSub$ = "Fnd_PrcBody_Tst"
'Dim mPrcBody$, mNmPrj_Nmm$, mNmPrc$, mFb$
'Dim mRslt As Boolean, mCase As Byte
'mCase = 4
'For mCase = 5 To 6
'    Select Case mCase
'    Case 1: mNmPrc = "zzGenDoc_FmtQry"
'    Case 2
'        mFb = "P:\WorkingDir\PgmObj\JMtcDb.mdb"
'        mNmPrj_Nmm = "JMtcDb.RunGenTbl"
'        mNmPrc = "qryGenTbl_TblCrtKey_Run"
'    Case 3
'        mFb = ""
'        mNmPrj_Nmm = "Fnd"
'        mNmPrc = "PrcBody_Tst"
'    Case 4
'        mFb = ""
'        mNmPrj_Nmm = "Fnd"
'        mNmPrc = "Prp"
'    Case 5
'        mFb = ""
'        mNmPrj_Nmm = "Acpt"
'        mNmPrc = "Dte_Tst"
'    Case 6
'        mFb = ""
'        mNmPrj_Nmm = "Acpt"
'        mNmPrc = "PkVal"
'    End Select
'Next
'Dim mAcs As Access.Application: If Cv_Acs_FmFb(mAcs, mFb) Then Stop: GoTo E
'mRslt = Fnd_PrcBody(mPrcBody, mNmPrj_Nmm, mNmPrc, mAcs)
'Shw_Dbg cSub, cMod, "mRslt,mNmPrj_Nmm,mNmPrc", mRslt, mNmPrj_Nmm, mNmPrc
'Debug.Print "------"
'Debug.Print Q_MrkUp(mPrcBody, "PrcBody")
'GoTo E
'E: Fnd_PrcBody_Tst = True
'X: If mFb <> "" Then Cls_CurDb mAcs
'End Function

'Function Fnd_Prp$(pNm$, pTypObj As AcObjectType, PrpNm$)
'Const cSub$ = "Fnd_Prp"
'On Error GoTo R
'Select Case pTypObj
'Case AcObjectType.acTable _
'     , AcObjectType.acReport _
'     , AcObjectType.acForm _
'     , AcObjectType.acMacro
'            Fnd_Prp = CurrentDb.Containers("Tables").Documents(pNm).Properties(PrpNm).Value
'Case AcObjectType.acQuery: Fnd_Prp = CurrentDb.QueryDefs(pNm).Properties(PrpNm).Value
'Case Else:  ss.A 1, "Invalid pTypObj": GoTo E
'End Select
'Exit Function
'R: ' ss.R
'E: ' ss.B cSub, cMod, "pNm,pTypObj,pNmPrd", pNm, ToStr_TypObj(pTypObj), PrpNm
'End Function
'Function Fnd_RecCnt_ByNmtq(oRecCnt&, Qry_or_Tbl_Nm$, Optional pLExpr$ = "", Optional pDb As database) As Boolean
'Const cSub$ = "Fnd_RecCnt_ByNmt"
'On Error GoTo R
'Dim mSql$: mSql = Fmt_Str("select count(*) from [{0}]{1}", Rmv_SqBkt(Qry_or_Tbl_Nm), SqlStrWhere(pLExpr))
'If Fnd_ValFmSql(oRecCnt, mSql, pDb) Then ss.A 1: GoTo E
'Exit Function
'R: ss.R
'E: Fnd_RecCnt_ByNmtq = True: ss.B cSub, cMod, "Qry_or_Tbl_Nm,pLExpr,pDb", Qry_or_Tbl_Nm, pLExpr, ToStr_Db(pDb)
'End Function

'Function Fnd_RecCnt_ByNmtq__Tst()
'Dim aa&
'Debug.Print Fnd_RecCnt_ByNmtq(aa, "mstAllBrand")
'Debug.Print aa
''
'Const cFb$ = "c:\aa.mdb"
'Dim mDb As database: If Crt_Db(mDb, cFb, True) Then Stop
'TblCrt_ByFldDclStr "aa", "aa Text 10", , , mDb) Then Stop
'Call mDb.Execute("Insert into aa values('abc')")
'Call mDb.Execute("Insert into aa values('abc')")
'If Fnd_RecCnt_ByNmtq(aa, "aa", , mDb) Then Stop
'Debug.Print aa
'Shw_DbgWin
'End Function

'
'Function Fnd_RecCnt_ByRs(oNRec%, pRs As DAO.Recordset) As Boolean
'Const cSub$ = "Fnd_RecCnt_ByRs"
'If IsNothing(pRs) Then Exit Function
'oNRec = 0
'On Error GoTo R
'With pRs
'    If .AbsolutePosition = -1 Then Exit Function
'    .MoveFirst
'    While Not .EOF
'        oNRec = oNRec + 1
'        .MoveNext
'    Wend
'    .MoveFirst
'End With
'Exit Function
'R: ss.R
'E: Fnd_RecCnt_ByRs = True: ss.B cSub, cMod, "pRs", ToStr_Rs_NmFld(pRs)
'End Function
'Function Fnd_ResStr(oStr$, pNmRes$, Optional pNmPrc_Nmm$ = "modResStr") As Boolean
'Const cSub$ = "Fnd_ResStr"
'If Fnd_PrcBody(oStr, pNmPrc_Nmm, pNmRes, , True) Then ss.A 1: GoTo E
'oStr = Cut_LastLin(Cut_FirstLin(Rmv_FirstChr(oStr)))
'Exit Function
'E: Fnd_ResStr = True: ss.B cSub, cMod, "pNmPrc_Nmm,pNmRes", pNmPrc_Nmm, pNmRes
'End Function

'Function Fnd_ResStr__Tst()
'Const cSub$ = "Fnd_ResStr_Tst"
'Dim mNmRes$, mStr$
'Dim mRslt As Boolean, mCase As Byte: mCase = 1
'Select Case mCase
'Case 1
'    mNmRes = "zzGenDoc_FmtQry"
'Case 2
'    mNmRes = "DtfTp"
'Case 3
'    mNmRes = "GenDoc_FmtMod"
'End Select
'mRslt = Fnd_ResStr(mStr, mNmRes)
'Shw_Dbg cSub, cMod, "mRslt,mNmRes,mStr", mRslt, mNmRes, mStr
'End Function


'Function Fnd_ResStr1__Tst()
'Dim mStr$
'If Fnd_ResStr(mStr, "Fnd_ResStr1_Tst", cMod) Then Stop
'Debug.Print mStr
'End Function

'Function Fnd_RgeCno_InRow(oRgeCno As tRgeCno, pWs As Worksheet, pRno&, Optional pCnoFm As Byte = 1, Optional pCnoTo As Byte = 255) As Boolean
''Aim: Looking for 'Beg' & 'End' in {pRno}
'Const cSub$ = "Fnd_RgeCno_InRow"
'With oRgeCno
'    .Fm = 0
'    .To = 0
'    Dim iCno As Byte: For iCno = pCnoFm To pCnoTo
'        If pWs.Cells(pRno, iCno).Value = "Beg" Then oRgeCno.Fm = iCno
'        If pWs.Cells(pRno, iCno).Value = "End" Then oRgeCno.To = iCno: Exit Function
'    Next
'End With
'ss.A 1, "Given row does not contain pair of Beg/End"
'E: Fnd_RgeCno_InRow = True: ss.B cSub, cMod, "pRno,pCnoFm,pCnoTo", pRno, pCnoFm, pCnoTo
'End Function
'Function Fnd_Rs_ByFilter(oRs As DAO.Recordset, pNmt$, pLExpr$) As Boolean
'Const cSub$ = "Fnd_Rs_ByFilter"
'On Error GoTo R
'Dim mSql$: mSql = Fmt_Str("Select * from {0} where {1}", Q_SqBkt(pNmt), pLExpr)
'Set oRs = CurrentDb.OpenRecordset(mSql)
'Exit Function
'R: ss.R
'E: Fnd_Rs_ByFilter = True: ss.B cSub, cMod, "pNmt,pLExpr", pNmt, pLExpr
'End Function
'Function Fnd_Rs_BySql(oRs As DAO.Recordset, Sql$, Optional pDb As database) As Boolean
'Const cSub$ = "Fnd_Rs_BySql"
'Dim mDb As database: Set mDb = DbNz(pDb)
'On Error GoTo R
'Set oRs = pDb.OpenRecordset(Sql)
'Exit Function
'R: ss.R
'E: Fnd_Rs_BySql = True: ss.B cSub, cMod, "Sql,pDb", Sql, ToStr_Db(pDb)
'End Function
'Function Fnd_SegFmCmd_2(oA1$, oA2$) As Boolean
'Const cSub$ = "Fnd_SegFmCmd_2"
'Dim mCmd$: mCmd = CtCommand()
'Dim mA$(): mA() = Split(mCmd, CtComma)
'If Siz_Ay(mA) <> 2 Then ss.A 1, "/Cmd is expected as {Nmrptsht},{NmSess} format.", , "mCmd", mCmd: GoTo E
'oA1 = mA(0)
'oA2 = mA(1)
'Exit Function
'E: Fnd_SegFmCmd_2 = True: ss.B cSub, cMod
'End Function
'Function Fnd_SegFmCmd_3(oA1$, oA2$, oA3$) As Boolean
'Const cSub$ = "Fnd_SegFmCmd_3"
'Dim mCmd$: mCmd = CtCommand()
'Dim mA$(): mA() = Split(mCmd, CtComma)
'If Siz_Ay(mA) <> 3 Then ss.A 1, "/Cmd is expected as {Nmrptsht},{NmSess}.{xxx} format.", , "mCmd", mCmd: GoTo E
'oA1 = mA(0)
'oA2 = mA(1)
'oA3 = mA(2)
'Exit Function
'E: Fnd_SegFmCmd_3 = True: ss.B cSub, cMod
'End Function
'Function Fnd_Sql_ByNmq(oSql$, QryNm$) As Boolean
'On Error GoTo R
'oSql = CurrentDb.QueryDefs(QryNm).Sql
'Exit Function
'R: ss.R
'E: Fnd_Sql_ByNmq = True
'End Function
'Function Fnd_Str_FmTxtFil(oS$, pFz$) As Boolean
'Const cSub$ = "Fnd_Str_FmTxtFil"
'On Error GoTo R
'Dim mFno As Byte: If Opn_Fil_ForInput(mFno, pFz) Then ss.A 1: GoTo E
'oS = ""
'While Not EOF(mFno)
'    Dim mL$: Line Input #mFno, mL
'    oS = Add_Str(oS, mL, vbCrLf)
'Wend
'Close #mFno
'Exit Function
'R: ss.R
'E: Fnd_Str_FmTxtFil = True: ss.B cSub, cMod, "pFz", pFz
'X:
'    Close #mFno
'End Function
'Function Fnd_Tbl(oTbl As DAO.TableDef, pNmt$) As Boolean
'Const cSub$ = "Fnd_Tbl"
'On Error GoTo R
'Set oTbl = CurrentDb.TableDefs(Rmv_SqBkt(pNmt))
'Exit Function
'R: ss.R
'E: Fnd_Tbl = True: ss.B cSub, cMod, "pNmt", pNmt
'End Function
'Public Function Fnd_TwoSq(oSqLeft As cSq, oSqTop As cSq, pWs As Worksheet) As Boolean
''Aim: Find 2Sq: {oSqLeft} & {oSqTop} by {pWs}, {pRno} & {pCno}.  {pRno} & {pCno} are bottom right corner of pane of the freezed window.
'Const cSub$ = "Fnd_TwoSq"
'If TypeName(oSqLeft) = "Nothing" Then Set oSqLeft = New cSq
'If TypeName(oSqTop) = "Nothing" Then Set oSqTop = New cSq
''Find the Freeze Cell
'Dim mFreezeRno&, mFreezeCno As Byte: If Fnd_FreezedCell(mFreezeRno, mFreezeCno, pWs) Then ss.A 1: GoTo E
''Detect mLasRno&, mLasCno
'Dim mLasRno&, mLasCno As Byte
'Dim mRge As Range: Set mRge = pWs.Cells.SpecialCells(xlCellTypeLastCell)
'mLasRno = mRge.Row
'mLasCno = mRge.Column
''Work From LasRno to pRno+1 to find first non-empty row so that oSqLeft is find
'Dim iRno&, iCno%, mIsEmpty As Boolean
'For iRno = mLasRno + 1 To mFreezeRno + 1 Step -1
'    mIsEmpty = True
'    For iCno = 1 To mFreezeCno
'        If Not IsEmpty(pWs.Cells(iRno, iCno).Value) Then mIsEmpty = False: Exit For
'    Next
'    If mIsEmpty Then Exit For
'Next
'With oSqLeft
'    .Cno1 = 1
'    .Cno2 = mFreezeCno
'    .Rno1 = mFreezeRno + 1
'    .Rno2 = iRno - 1
'End With
''Work From LasCno to pCno+1 to find first non-empty column so that oSqTop is find
'For iCno = mLasCno + 1 To mFreezeCno + 1 Step -1
'    mIsEmpty = True
'    For iRno = 1 To mFreezeRno
'        If Not IsEmpty(pWs.Cells(iRno, iCno).Value) Then mIsEmpty = False: Exit For
'    Next
'    If mIsEmpty Then Exit For
'Next
'With oSqTop
'    .Cno1 = mFreezeCno + 1
'    .Cno2 = iCno - 1
'    .Rno1 = 1
'    .Rno2 = mFreezeRno
'End With
'Exit Function
'E: Fnd_TwoSq = True: ss.B cSub, cMod, "pWs", ToStr_Ws(pWs)
'End Function
'Private Function Fnd_TwoSq__Tst()
''Debug.Print Application.Workbooks.Count
''Debug.Print Application.Workbooks(1).FullName
''Dim mWs As Worksheet: Set mWs = Application.Workbooks(1).Sheets("Total")
''Dim mSqLeft As cSq, mSqTop As cSq
''If TwoSq(mSqLeft, mSqTop, mWs) Then Stop
''Debug.Print "mSqLeft=" & mSqLeft.ToStr
''Debug.Print "mSqTop=" & mSqTop.ToStr
'End Function
'Function Fnd_TypPrmRpt(oTypPrmRpt As tRpt, pNmRptSht$) As Boolean
'Const cSub$ = "Fnd_TypPrmRpt"
'On Error GoTo R
'Dim mRs As DAO.Recordset: If Opn_Rs(mRs, "Select * from tblRpt where Nmrptsht='" & pNmRptSht & CtSngQ) Then ss.A 1: GoTo E
'With mRs
'    If .AbsolutePosition = -1 Then ss.A 1, "Given report not define in tblRpt": GoTo E
'    oTypPrmRpt.NmRpt = !NmRpt
'    oTypPrmRpt.FmtStr_FnTo = Nz(!FmtStr_FnTo, "")
'    oTypPrmRpt.QryPrm = Nz(!QryPrm, "")
'    oTypPrmRpt.LnwsRmv = Nz(!LnwsRmv, "")
'    oTypPrmRpt.HidePfLst_ThisNmSess = Nz(!HidePfLst_ThisNmSess, "")
'    oTypPrmRpt.HidePfLst_ThisSess = Nz(!HidePfLst_ThisSess, "")
'    oTypPrmRpt.HidePfLst_OtherSess = Nz(!HidePfLst_OtherSess, "")
'    oTypPrmRpt.NmDta = Nz(!NmDta, "")
'    oTypPrmRpt.EachSql = Nz(!EachSql, "")
'    oTypPrmRpt.EachNmFld = Nz(!EachNmFld, "")
'    oTypPrmRpt.EachLnwsRmv = Nz(!EachLnwsRmv, "")
'    oTypPrmRpt.EachHidePfLst_ThisSess = Nz(!EachHidePfLst_ThisSess, "")
'    oTypPrmRpt.EachHidePfLst_OtherSess = Nz(!EachHidePfLst_OtherSess, "")
'End With
'GoTo X
'R: ss.R
'E: Fnd_TypPrmRpt = True: ss.B cSub, cMod, "pNmRptSht", pNmRptSht
'X: RsCls mRs
'End Function
'Function Fnd_LvFmRs_Of1Rec(oLv$, pRs As DAO.Recordset, FnStr$, Optional pSepChr$ = CtCommaSpc) As Boolean
'Const cSub$ = "Fnd_LvFmRs_Of1Rec"
'On Error GoTo R
'Dim mAnFld$(): mAnFld = Split(FnStr, CtComma)
'Dim N%: N = Siz_Ay(mAnFld)
'Dim mV
'oLv = ""
'Dim J%: For J = 0 To N - 1
'    If Fnd_FldVal_ByFld(mV, pRs.Fields(J)) Then ss.A 1, "One of the field cannot Get Fld Val", , "J", J: GoTo E
'    oLv = Add_Str(oLv, Q_V(mV), pSepChr)
'Next
'Exit Function
'R: ss.R
'E: Fnd_LvFmRs_Of1Rec = True: ss.B cSub, cMod, "pRs,FnStr,pSepChr", ToStr_Rs_NmFld(pRs), FnStr, pSepChr
'End Function
'Function Fnd_LvFmRs(oLv$, pRs As DAO.Recordset, Optional pNmFld$ = "", Optional pQ$ = "", Optional pSepChr$ = CtComma) As Boolean
''Aim: FInd {oLv} from all record of first field  <pRs>.<pNmFld> of each record in {pRs} into {oLv}
'Const cSub$ = "Fnd_LvFmRs"
'oLv = ""
'On Error GoTo R
'Dim mAyV(): If RsCol(mAyV, pRs, pNmFld) Then ss.A 1: GoTo E
'oLv = Join_AyV(mAyV, pQ, pSepChr)
'Exit Function
'R: ss.R
'E: Fnd_LvFmRs = True: ss.B cSub, cMod, "pRs,pNmFld", ToStr_Rs(pRs), pNmFld
'End Function

'Function Fnd_LvFmRs__Tst()
'Const cSub$ = "Fnd_LvFmRs"
'Dim mNmt$, mNmFld$, mLv$, mCase As Byte
'
'Shw_Dbg cSub, cMod
'For mCase = 1 To 1
'    Select Case mCase
'    Case 1: mNmt = "mstBrand": mNmFld = "BrandId"
'    Case 2
'    Case 3
'    End Select
'    Dim mRs As DAO.Recordset: Set mRs = CurrentDb.TableDefs(mNmt).OpenRecordset
'    If Fnd_LvFmRs(mLv, mRs, mNmFld) Then Stop
'    mRs.Close
'    Debug.Print ToStr_LpAp(vbLf, "mCase,mNmt,mNmFld,mLv", mCase, mNmt, mNmFld, mLv)
'    Debug.Print "----"
'Next
'End Function

'Function Fnd_ValFmSql(oVal, Sql$ _
'    , Optional pDb As database _
'    ) As Boolean
''Aim: a value from a 'scalar' {Sql} in {pDb}
'Const cSub$ = "Fnd_ValFmSql"
'On Error GoTo R
'Dim mDb As database: Set mDb = DbNz(pDb)
'Dim mRs As DAO.Recordset: If Opn_Rs(mRs, Sql, pDb) Then ss.A 1: GoTo E
'With mRs
'    If .AbsolutePosition = -1 Then ss.A 1, "no record": GoTo E
'    oVal = .Fields(0).Value
'End With
'GoTo X
'R: ss.R
'E: Fnd_ValFmSql = True: ss.B cSub, cMod, "Sql,pDb", Sql, ToStr_Db(pDb)
'X: RsCls mRs
'End Function

'Function Fnd_ValFmSql__Tst()
'Dim mSql$: mSql = "Select * from [#OldPgm]"
'Dim mA$: If Fnd_ValFmSql(mA, mSql) Then Stop: GoTo E
'Debug.Print mA
'Exit Function
'E: Fnd_ValFmSql_Tst = True
'End Function

'Function Fnd_ValFmSql2(oV1, oV2, Sql$ _
'    , Optional pDb As database _
'    ) As Boolean
'Const cSub$ = "Fnd_ValFmSql2"
''Aim: Find first 2 values from the first record of {Sql}
'On Error GoTo R
'Dim mDb As database: Set mDb = DbNz(pDb)
'Dim mRs As DAO.Recordset: If Opn_Rs(mRs, Sql, pDb) Then ss.A 1: GoTo E
'With mRs
'    If .AbsolutePosition = -1 Then GoTo E
'    oV1 = .Fields(0).Value
'    oV2 = .Fields(1).Value
'End With
'GoTo X
'R: ss.R
'E: Fnd_ValFmSql2 = True: ss.B cSub, cMod, "Sql,pDb", Sql, ToStr_Db(pDb)
'X:
'    RsCls mRs
'End Function
'Function Fnd_ValFmSql3(oV1, oV2, oV3, Sql$ _
'    , Optional pDb As database _
'    ) As Boolean
'Const cSub$ = "Fnd_ValFmSql3"
''Aim: Find first 3 values from the first record of {Sql}
'On Error GoTo R
'Dim mDb As database: Set mDb = DbNz(pDb)
'Dim mRs As DAO.Recordset: If Opn_Rs(mRs, Sql, pDb) Then ss.A 1: GoTo E
'With mRs
'    If .AbsolutePosition = -1 Then GoTo E
'    oV1 = .Fields(0).Value
'    oV2 = .Fields(1).Value
'    oV3 = .Fields(2).Value
'End With
'GoTo X
'R: ss.R
'E: Fnd_ValFmSql3 = True: ss.B cSub, cMod, "Sql,pDb", Sql, ToStr_Db(pDb)
'X:
'    RsCls mRs
'End Function
'Function Fnd_ValFmSql4(oV1, oV2, oV3, oV4, Sql$ _
'    , Optional pDb As database _
'    ) As Boolean
'Const cSub$ = "Fnd_ValFmSql4"
''Aim: Find first 4 values from the first record of {Sql}
'On Error GoTo R
'Dim mDb As database: Set mDb = DbNz(pDb)
'Dim mRs As DAO.Recordset: If Opn_Rs(mRs, Sql, pDb) Then ss.A 1: GoTo E
'With mRs
'    If .AbsolutePosition = -1 Then GoTo E
'    oV1 = .Fields(0).Value
'    oV2 = .Fields(1).Value
'    oV3 = .Fields(2).Value
'    oV4 = .Fields(3).Value
'End With
'GoTo X
'R: ss.R
'E: Fnd_ValFmSql4 = True: ss.B cSub, cMod, "Sql,pDb", Sql, ToStr_Db(pDb)
'X:
'    RsCls mRs
'End Function
'Function Fnd_ValFmSql5(oV1, oV2, oV3, oV4, oV5, Sql$ _
'    , Optional pDb As database _
'    ) As Boolean
'Const cSub$ = "Fnd_ValFmSql5"
''Aim: Find first 5 values from the first record of {Sql}
'On Error GoTo R
'Dim mDb As database: Set mDb = DbNz(pDb)
'Dim mRs As DAO.Recordset: If Opn_Rs(mRs, Sql, pDb) Then ss.A 1: GoTo E
'With mRs
'    If .AbsolutePosition = -1 Then GoTo E
'    oV1 = .Fields(0).Value
'    oV2 = .Fields(1).Value
'    oV3 = .Fields(2).Value
'    oV4 = .Fields(3).Value
'    oV5 = .Fields(4).Value
'End With
'GoTo X
'R: ss.R
'E: Fnd_ValFmSql5 = True: ss.B cSub, cMod, "Sql,pDb", Sql, ToStr_Db(pDb)
'X:
'    RsCls mRs
'End Function
'Function Fnd_ValFmTbl_ByWhere(oVal, pNmt$, pNmFldRet$, pWhere$) As Boolean
'Const cSub$ = "Fnd_ValFmTbl_ByWhere"
'On Error GoTo R
'Dim mSql$: mSql = Fmt_Str("Select {0} from {1} where {2}", pNmFldRet, Q_SqBkt(pNmt), pWhere)
'Dim mRs As DAO.Recordset: If Opn_Rs(mRs, mSql) Then ss.A 1: GoTo E
'With mRs
'    If .AbsolutePosition = -1 Then ss.A 2, "No record found by given mWhere in table": GoTo E
'    oVal = .Fields(pNmFldRet)
'End With
'GoTo X
'R: ss.R
'E: Fnd_ValFmTbl_ByWhere = True: ss.B cSub, cMod, "pNmt,pNmFldRet,pWhere", pNmt, pNmFldRet, pWhere
'X: RsCls mRs
'End Function
'Function Fnd_VbCmp(oVbCmp As VBComponent, pVBPrj As vbproject, pNmCmp$) As Boolean
'Const cSub$ = "Fnd_VBCmp"
'On Error GoTo R
'Set oVbCmp = pVBPrj.VBComponents(pNmCmp)
'Exit Function
'R: Fnd_VbCmp = True
'End Function

'Function Fnd_VbCmp__Tst()
'Dim mPrj As vbproject: If Fnd_Prj(mPrj, "jj") Then Stop: GoTo E
'Dim mVbCmp As VBComponent: If Fnd_VbCmp(mVbCmp, mPrj, "Form_frmWaitFor") Then Stop: GoTo E
'Stop
'Exit Function
'E: Fnd_VbCmp_Tst = True
'End Function

'Function Fnd_VbCmp_FmWs(oVbCmp As VBIDE.VBComponent, pWs As Worksheet) As Boolean
'Const cSub$ = "Fnd_VBCmp_FmWs"
'On Error GoTo R
'Dim mNmWs$: mNmWs = pWs.Name
'For Each oVbCmp In pWs.Application.Vbe.ActiveVBProject.VBComponents
'    If oVbCmp.Type = vbext_ct_Document Then
'        If oVbCmp.Properties("Name").Value = pWs.Name Then Exit Function
'    End If
'Next
'ss.A 1, "VBCmp not find for ws"
'GoTo E
'R: ss.R
'E: Fnd_VbCmp_FmWs = True: ss.C cSub, cMod, "VBCmp not find for ws", "pWs", ToStr_Ws(pWs)
'End Function
'Function Fnd_Md_ByNm(oMd As CodeModule, pMod$ _
'    , Optional pAcs As Access.Application = Nothing _
'    ) As Boolean
'Const cSub$ = "Fnd_Md_ByMogd"
'Dim mNmPrj$, mNmm$: If Brk_Str_Both(mNmPrj, mNmm, pMod, ".") Then ss.A 1: GoTo E
'Dim mPrj As vbproject: If Fnd_Prj(mPrj, mNmPrj, pAcs) Then ss.A 2: GoTo E
'If Fnd_Md(oMd, mPrj, mNmm) Then ss.A 3: GoTo E
'Exit Function
'E: Fnd_Md_ByNm = True: ss.B cSub, cMod, "pMod,pAcs", pMod, ToStr_Acs(pAcs)
'End Function

'Function Fnd_Md_ByNm__Tst()
'Dim mMd As CodeModule: If Fnd_Md_ByNm(mMd, "cSq") Then Stop
'Debug.Print ToStr_Md(mMd)
'Stop
'End Function

'Function Fnd_Md(oMd As CodeModule, pPrj As vbproject, pNmm$ _
'    ) As Boolean
'Const cSub$ = "Fnd_Md"
'On Error GoTo R
'Dim iCmp As VBComponent
'Set iCmp = pPrj.VBComponents(pNmm)
'Set oMd = iCmp.CodeModule
'Exit Function
'GoTo E
'R: ss.R
'E: Fnd_Md = True: ss.C cSub, cMod, "pPrj,pNmm", ToStr_Prj(pPrj), pNmm
'End Function
'Function Fnd_PrcRgeRno(oRgeRno As tRgeRno, pMod$, pNmPrc$) As Boolean
''Aim: Find line range {oRgeRno} of {pNmPrc} in {pMd}
'Const cSub$ = "Fnd_PrcRgeRno"
'Dim mMd As CodeModule: If Fnd_Md_ByNm(mMd, "xFnd") Then ss.A 1: GoTo E
'Fnd_PrcRgeRno = Fnd_PrcRgeRno_ByMd(oRgeRno, mMd, pNmPrc)
'Exit Function
'E: Fnd_PrcRgeRno = True: ss.A cSub, cMod, "pMod,pNmPrc", pMod, pNmPrc
'End Function

'Function Fnd_PrcRgeRno__Tst()
'Shw_DbgWin
'Dim mRgeRno As tRgeRno
'If Fnd_PrcRgeRno(mRgeRno, "xFnd", "Fnd_PrcRgeRno_Tst") Then Stop
'Debug.Print mRgeRno.Fm, mRgeRno.To
'If Fnd_PrcRgeRno(mRgeRno, "xFnd", "Fnd_PrcRgeRno_ByMd") Then Stop
'Debug.Print mRgeRno.Fm, mRgeRno.To
'End Function

'Function Fnd_PrcRgeRno_ByMd(oRgeRno As tRgeRno, pMd As CodeModule, pNmPrc$) As Boolean
''Aim: Find line range {oRgeRno} of {pMod}.{pNmPrc}
'Const cSub$ = "Fnd_PrcRgeRno_ByMd"
'On Error GoTo R
'Dim mAnPrc_LinBeg_LinEnd$(): If Fnd_AnPrc_ByMd(mAnPrc_LinBeg_LinEnd, pMd, pNmPrc, , True) Then ss.A 1: GoTo E
'Dim mN%: mN = Siz_Ay(mAnPrc_LinBeg_LinEnd)
'If mN = 0 Then oRgeRno.Fm = 0: oRgeRno.To = 0: Exit Function
'If mN > 1 Then ss.A 2, "Return mAnPrc_LinBeg_LinEnd should be one element", , "mAnPrc_LinBeg_LinEnd", ToStr_Ays(mAnPrc_LinBeg_LinEnd): GoTo E
'Dim mNmPrc$, mLinBeg&, mLinEnd&: If Brk_Str_To3Seg(mNmPrc, oRgeRno.Fm, oRgeRno.To, mAnPrc_LinBeg_LinEnd(0), ":") Then ss.A 2: GoTo E
'Exit Function
'R: ss.R
'E: Fnd_PrcRgeRno_ByMd = True: ss.B cSub, cMod, "pMd,pNmPrc", ToStr_Md(pMd), pNmPrc
'End Function
'Function Fnd_AnPrc_ByMd(oAnPrc_LinBeg_LinEnd$(), pMd As CodeModule _
'    , Optional pLikNmPrc$ = "*" _
'    , Optional pSrt As Boolean = False _
'    , Optional pWithLinNo As Boolean = False _
'    , Optional pBodyOnly As Boolean = False _
'    ) As Boolean
''Aim: All lines begin and end being empty line or start with #.
'Const cSub$ = "Fnd_AnPrc_ByMd"
'On Error GoTo R
'Clr_Ays oAnPrc_LinBeg_LinEnd
'With pMd
'    Dim iLinNo&: iLinNo = .CountOfDeclarationLines + 1
'    Dim iNmPrc$, iBeg&, iEnd&
'    While iLinNo < .CountOfLines
'        Dim mVBExt_pk_Proc As vbext_ProcKind
'        iNmPrc = .ProcOfLine(iLinNo, mVBExt_pk_Proc)
'        If iNmPrc Like pLikNmPrc Then
'
'            If pWithLinNo Then
'                iBeg = .ProcStartLine(iNmPrc, mVBExt_pk_Proc)
'                iEnd = .ProcCountLines(iNmPrc, mVBExt_pk_Proc) + iBeg - 1
'                Dim mBeg&, mEnd&, mL$, mA$
'                mBeg = iBeg
'                mEnd = iEnd
'                For mBeg = iBeg To iEnd
'                    mL = Trim(.Lines(mBeg, 1)): mA = Left(mL, 1)
'                    If mL <> "" And mA <> CtSngQ And mA <> "#" Then Exit For
'                Next
'
'                For mEnd = iEnd To mBeg Step -1
'                    mL = Trim(.Lines(mEnd, 1)): mA = Left(mL, 1)
'                    If mL <> "" And mA <> CtSngQ And mA <> "#" Then Exit For
'                Next
'                If Not pBodyOnly Then
'                    mL = Trim(.Lines(mBeg - 1, 1)): mA = Left(mL, 3)
'                    If mA = "#If" Then mBeg = mBeg - 1
'                    mL = Trim(.Lines(mEnd + 1, 1)): mA = Left(mL, 7)
'                    If mA = "#End If" Then mEnd = mEnd + 1
'                End If
'                Add_AyEle oAnPrc_LinBeg_LinEnd, iNmPrc & ":" & mBeg & ":" & mEnd
'            Else
'                Add_AyEle oAnPrc_LinBeg_LinEnd, iNmPrc
'            End If
'        End If
'        iLinNo = iLinNo + .ProcCountLines(iNmPrc, mVBExt_pk_Proc)
'    Wend
'End With
'If pSrt Then If Srt_Ay(oAnPrc_LinBeg_LinEnd, oAnPrc_LinBeg_LinEnd) Then ss.A 1: GoTo E
'Exit Function
'R: ss.R
'E: Fnd_AnPrc_ByMd = True: ss.C cSub, cMod, "pMd,pLikNmPrc,pSrt", ToStr_Md(pMd), pLikNmPrc, pSrt
'End Function

'Function Fnd_AnPrc__Tst()
'Const cSub$ = "Fnd_AnPrc_Tst"
'Dim J%
'Dim mPrj As vbproject:
'Dim mAnm$(), mAnPrj$(), mAnPrc$(), mMd As CodeModule
'Dim mCase As Byte
'mCase = 2
'Select Case mCase
'Case 1
'    If Fnd_AnPrj(mAnPrj) Then Stop: GoTo E
'    For J = 0 To Siz_Ay(mAnPrj) - 1
'        If Fnd_Prj(mPrj, mAnPrj(J)) Then Stop: GoTo E
'        If Fnd_Anm_ByPrj(mAnm, mPrj) Then Stop: GoTo E
'        Dim I%
'        For I = 0 To Siz_Ay(mAnm) - 1
'            If Fnd_Md(mMd, mPrj, mAnm(I)) Then Stop: GoTo E
'            If Fnd_AnPrc_ByMd(mAnPrc, mMd) Then Stop: GoTo E
'            Debug.Print mAnPrj(J) & "." & mAnm(I) & ": " & ToStr_Ays(mAnPrc)
'        Next
'    Next
'Case 2
'    Dim mLikPrc$:   mLikPrc = "qry*"
'    Dim mNmPrj$:    mNmPrj = "JMtcDb"
'    Dim mNmm$:      mNmm = "RunGentTbl"
'    Dim mFbSrc$:    mFbSrc = "P:\WorkingDir\PgmObj\JMtcDb.mdb"
'    Dim mAcs As Access.Application: If Cv_Acs_FmFb(mAcs, mFbSrc) Then Stop: GoTo E
'    If Fnd_Prj(mPrj, mNmPrj, mAcs) Then Stop: GoTo E
'    If Fnd_Md(mMd, mPrj, mNmm) Then Stop: GoTo E
'    If Fnd_AnPrc_ByMd(mAnPrc, mMd, mLikPrc, , True) Then Stop
'    Debug.Print ToStr_Ays(mAnPrc, , vbLf)
'End Select
'Exit Function
'R: ss.R
'E: Fnd_AnPrc_Tst = True: ss.B cSub, cMod
'End Function

'Function Fnd_Prj(oPrj As vbproject, pNmPrj, Optional pApp As Application) As Boolean
'Const cSub$ = "Fnd_Prj"
'On Error GoTo R
'Dim mApp As Application: Set mApp = Cv_App(pApp)
'Set oPrj = mApp.Vbe.VBProjects(pNmPrj)
'Exit Function
'R: ss.R
'E: Fnd_Prj = True: ss.C cSub, cMod, "pNmPrj,pAcs", pNmPrj, ToStr_App(pApp)
'End Function
'
'

