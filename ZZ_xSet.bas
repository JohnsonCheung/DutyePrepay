Attribute VB_Name = "ZZ_xSet"
'
'Option Compare Text
'Option Explicit
'Option Base 0
'Const cMod$ = cLib & ".xSet"
'Dim x_AySilent() As Boolean
'Dim x_AyNoLog() As Boolean
'Function Set_Import_AtA1(pFx$) As Boolean
''Aim: each ws in {pFx} set a1 as "Import:{WsNm}"
'Const cSub$ = "Set_Import_AtA1"
'On Error GoTo R
'Dim mWb As Workbook: If Opn_Wb_RW(mWb, pFx) Then ss.A 1: GoTo E
'Dim iWs As Worksheet
'For Each iWs In mWb.Sheets
'    iWs.Range("A1").Value = "Import:" & iWs.Name
'Next
'If Cls_Wb(mWb, True) Then ss.A 2: GoTo E
'Exit Function
'R: ss.R
'E: Set_Import_AtA1 = True: ss.B cSub, cMod, "pFx", pFx
'End Function
'
'Function Set_Import_AtA1__Tst()
'If Set_Import_AtA1("P:\AppDef_Meta\MetaDb.xls") Then Stop
'End Function
'
'Function Set_Lm_FmSql(oLm$, Sql$ _
'    , Optional pNmFld0$ = "" _
'    , Optional pNmFld1$ = "" _
'    , Optional pBrkChr$ = "=" _
'    , Optional pSepChr$ = vbCrLf) As Boolean
''Aim: Build {oLm} from 2 fields ({pNmFld1} & {pNmFld2}) of {pRs}
'Const cSub$ = "Set_Lm_FmSql"
'Dim mRs As DAO.Recordset: If Opn_Rs(mRs, Sql) Then ss.A 1: GoTo E
'If RsToStr(oLm, mRs, pNmFld0, pNmFld1, pBrkChr, pSepChr) Then ss.A 2: GoTo E
'GoTo X
'R: ss.R
'E: Set_Lm_FmSql = True: ss.B cSub, cMod, "Sql,pNmFld0,pNmFld1,pBrkChr,pSepChr", Sql, pNmFld0, pNmFld1, pBrkChr, pSepChr
'X: RsCls mRs
'End Function
'Function RsSel(Rs As DAO.Recordset, FnStr$) As Dictionary
''Aim: Build {oLv} by {FnStr} in {pRs}
'Stop
''Dim mAnFld_Lcl$(), mAnFld_Host$(): If Brk_Lm_To2Ay(mAnFld_Lcl, mAnFld_Host, FnStr) Then ss.A 1: GoTo E
''Dim N%: N = Sz(mAnFld_Lcl)
''With Rs
''    Dim J%, mA$
''    If pIsNoNm Then
''        For J = 0 To N - 1
''            oLv = Add_Str(oLv, Q_V(.Fields(mAnFld_Lcl(J)).Value), pSep$)
''        Next
''    Else
''        For J = 0 To N - 1
''            If Join_NmV(mA, mAnFld_Host(J), .Fields(mAnFld_Lcl(J)).Value, pBrk) Then ss.A 1: GoTo E
''            oLv = Add_Str(oLv, mA, pSep$)
''        Next
''    End If
''End With
'End Function
'
'Function RsToStr__Tst()
'TblCrt_ByFldDclStr "#Tmp", "Itm Text 10,N Text 50,X Text 50"
'If Run_Sql("Insert into [#Tmp] values ('Tbl','1,2,3',',x,xx,xxx')") Then Stop
'Dim mLm$: If Set_Lm_ByTbl(mLm, "#Tmp") Then Stop
'Debug.Print mLm
'End Function
'Function RsToStr(oLm$, pRs As DAO.Recordset _
'    , Optional pNmFld0$ = "" _
'    , Optional pNmFld1$ = "" _
'    , Optional pBrkChr$ = "=" _
'    , Optional pSepChr$ = vbCrLf) As Boolean
''Aim: Build {oLm} from all records in {pRs} which have 2 fields {pNmFld1} & {pNmFld2}
'Const cSub$ = "RsToStr"
'On Error GoTo R
'Dim mNmFld0$: mNmFld0 = NonBlank(pNmFld0, pRs.Fields(0).Name)
'Dim mNmFld1$: mNmFld1 = NonBlank(pNmFld1, pRs.Fields(1).Name)
'oLm = ""
'With pRs
'    While Not .EOF
'        oLm = Add_Str(oLm, .Fields(mNmFld0).Value & pBrkChr & .Fields(mNmFld1).Value, pSepChr)
'        .MoveNext
'    Wend
'End With
'Exit Function
'R: ss.R
'E: RsToStr = True: ss.B cSub, cMod, "pRs,pNmFld0,pNmFld1,pBrkChr,pSepChr", ToStr_Flds(pRs.Fields), pNmFld0, pNmFld1, pBrkChr, pSepChr
'End Function
'
'Sub Set_TBar_Toggle()
'Dim mWs As Worksheet: Set mWs = Excel.Application.ActiveSheet
'If IsNothing(mWs) Then Exit Sub
'Dim iOLEObj As Excel.OLEObject
'For Each iOLEObj In mWs.OLEObjects
'    If TypeName(iOLEObj.Object) = "ToolBar" Then iOLEObj.Visible = Not iOLEObj.Visible
'Next
'End Sub
'
'Function Set_Ws_ByAyV(pWs As Worksheet, pRno&, pCno As Byte, pIsDown As Boolean, pAyV()) As Boolean
'Set_Ws_ByAyV = Set_Ws_ByVayv(pWs, pRno, pCno, pIsDown, CVar(pAyV))
'End Function
'Function Set_Ws_ByVayv(pWs As Worksheet, pRno&, pCno As Byte, pIsDown As Boolean, pVayv) As Boolean
''Aim: Set {pRno} in {pWs} by {pAp}
'Const cSub$ = "Set_Ws_ByVayv"
'On Error GoTo R
'Dim mAyV(): mAyV = pVayv
'Dim J%, N%: N = Sz(mAyV)
'With pWs
'    If pIsDown Then
'        For J = 0 To N - 1
'            .Cells(pRno + J, pCno).Value = mAyV(J)
'        Next
'        Exit Function
'    End If
'    For J = 0 To N - 1
'        .Cells(pRno, pCno + J).Value = mAyV(J)
'    Next
'End With
'Exit Function
'R: ss.R
'E: Set_Ws_ByVayv = True: ss.B cSub, cMod, "pWs,pRno,pCno,pIsDown,Vayv,", ToStr_Ws(pWs), pRno, pCno, pIsDown, ToStr_Vayv(pVayv)
'End Function
'Function Set_Ws_ByAyPrm(pWs As Worksheet, pRno&, pCno As Byte, pIsDown As Boolean, ParamArray pAp()) As Boolean
''Aim: Set {pRno} in {pWs} by {pAp}
'Set_Ws_ByAyPrm = Set_Ws_ByVayv(pWs, pRno, pCno, pIsDown, CVar(pAp))
'End Function
'Public Sub Set_Ws_CmbBox(pWs As Excel.Worksheet, pPfx As String, pCtlCnt As Byte, pPrp As String, V)
''Aim: Assume there are pCtlCnt comboxbox control object in the pWs with name XXX01, ... XXXnn, where XXX is pPfx, nn is pCtlCnt
''     It is required to set the property pPrp for each of the control by the value V
'Dim J As Byte
'For J = 1 To 20
'    Select Case pPrp
'    Case "ListFillRange":    pWs.OLEObjects(pPfx & Format(J, "00")).ListFillRange = V
'    Case "PrintObject":      pWs.OLEObjects(pPfx & Format(J, "00")).PrintObject = V
'    Case "Height":           pWs.OLEObjects(pPfx & Format(J, "00")).Height = V
'    Case "Height":           pWs.OLEObjects(pPfx & Format(J, "00")).Height = V
'    Case "ListRows":         pWs.OLEObjects(pPfx & Format(J, "00")).ListRows = V
'    Case Else
'    Stop
'    End Select
'Next
'End Sub
'
'Function Set_Fld_ToAuto(pNmt$, pNmFld$) As Boolean
'Const cSub$ = "Set_Fld_ToAuto"
'On Error GoTo R
'Dim mFldAtr&: mFldAtr = CurrentDb.TableDefs(pNmt).Fields(pNmFld).Attributes
'CurrentDb.TableDefs(pNmt).Fields(pNmFld).Attributes = mFldAtr Or DAO.FieldAttributeEnum.dbAutoIncrField
'Exit Function
'R: ss.R
'E: Set_Fld_ToAuto = True: ss.B cSub, cMod, "pNmt,pNmFld", pNmt, pNmFld
'End Function
'
'Function Set_Rs_ByLpVv(ORs As Recordset, Lp$, Av) As Boolean
''See DicSetRs
'End Function
'Sub DicSetRs__Tst()
'If Dlt_Tbl("xx") Then Stop
'If Run_Sql("Create table xx (aa Long, bb Integer, cc Date)") Then Stop
'Dim mRs As DAO.Recordset
'Set mRs = CurrentDb.TableDefs("xx").OpenRecordset
'mRs.AddNew
''DicSetRs mRs, "aa,bb,cc", "13", 12, "2007/12/31") Then Stop ' Should have NO error
'mRs.Update
'
'mRs.AddNew
''If DicSetRs(mRs, "aa,bb,cc", 13, 12, #1/1/2007#) Then Stop ' Should have NO error
'mRs.Update
'
'mRs.AddNew
''If DicSetRs(mRs, "aa,bb,cc", "13a", 12, #1/1/2007#) Then Stop ' Should have error
'mRs.Update
'mRs.Close
'DoCmd.OpenTable ("xx")
'End Sub
'
'Sub DicSetRs(Dic As Dictionary, ORs As DAO.Recordset)
''Aim: Set {oRs} by {pLnFld} & {pAyV}.  Assume oRs is already .AddNew or .Edit
'Const cSub$ = "Set_Rs_ByLpVv"
'
'Dim J%, mAnFld$(): 'mAnFld = Split(pLnFld, cComma)
'Dim mNmFld$, mAyV()
''mAyV = pVayv
'With ORs
'    For J = 0 To UBound(mAnFld$)
'        mNmFld = Trim(mAnFld(J))
'        .Fields(mNmFld).Value = mAyV(J)
'    Next
'End With
'End Sub
'
'Function Set_ChdLnk(pFrm As Access.Form, pLnChd$, pMst$, pChd$) As Boolean
'Const cSub$ = "Set_ChdLnk"
'On Error GoTo R
'Dim mAnChd$(): mAnChd = Split(pLnChd, CtComma)
'Dim J%: For J = 0 To Sz(mAnChd) - 1
'    Dim mSubFrm As SubForm: If Fnd_Ctl(mSubFrm, pFrm, mAnChd(J)) Then GoTo Nxt
'    With mSubFrm
'        .LinkMasterFields = pMst
'        .LinkChildFields = pChd
'    End With
'Nxt:
'Next
'Exit Function
'R: ss.R
'
'E:
': ss.B cSub, cMod, "pFrm,pLnChd,pChd,pMst", ToStr_Frm(pFrm), , pLnChd, pChd, pMst
'    Set_ChdLnk = True
'End Function
'Function Set_EnableEdt(pFrm As Access.Form, pEnable As Boolean) As Boolean
'Const cSub$ = "Set_EnableEdt"
'Dim iCtl As Access.Control: For Each iCtl In pFrm.Controls
'    If iCtl.Tag = "Edt" Then
'        Dim mNmTyp$: mNmTyp = TypeName(iCtl)
'        Select Case mNmTyp
'        Case "TextBox":  If Set_EnableTBox(iCtl, pEnable) Then ss.A 1: GoTo E
'        Case "Check":    If Set_EnableChkB(iCtl, pEnable) Then ss.A 2: GoTo E
'        Case "ComboBox": If Set_EnableCBox(iCtl, pEnable) Then ss.A 3: GoTo E
'        End Select
'    End If
'Next
'E: Set_EnableEdt = True: ss.B cSub, cMod, "pFrm,pEnable", ToStr_Frm(pFrm), pEnable
'End Function
'Function Set_EnableChkB(pChkB As Access.CheckBox, pEnable As Boolean) As Boolean
'Const cSub$ = "Set_EnableChkB"
'pChkB.Enabled = pEnable
'pChkB.BorderColor = IIf(pEnable, 65280, 13209)
'On Error Resume Next
'End Function
'Function Set_EnableCBox(pCBox As Access.ComboBox, pEnable As Boolean) As Boolean
'Const cSub$ = "Set_EnableCBox"
'pCBox.Enabled = pEnable
'pCBox.ForeColor = IIf(pEnable, 0, 255)
'On Error Resume Next
'End Function
'Function Set_EnableTBox(pTBox As Access.TextBox, pEnable As Boolean) As Boolean
'Const cSub$ = "Set_EnableTBox"
'pTBox.Enabled = pEnable
'pTBox.ForeColor = IIf(pEnable, 0, 255)
'On Error Resume Next
'End Function
'Function Set_CmdBtnSte(pNmCmdBar$, pLnBtn$, pBtnSte As MsoButtonState) As Boolean
'Const cSub$ = "Set_CmdBtn"
'On Error GoTo R
'Dim mCmdBar As CommandBar: Set mCmdBar = Application.CommandBars(pNmCmdBar)
'Dim mAnBtn$(): mAnBtn = Split(pLnBtn, CtComma)
'Dim mEnabled As Boolean: mEnabled = (pBtnSte <> msoButtonDown)
'With mCmdBar
'    Dim J%: For J = 0 To Sz(mAnBtn) - 1
'        With .Controls(mAnBtn(J))
'            .State = pBtnSte
'            .Enabled = mEnabled
'        End With
'    Next
'End With
'Exit Function
'R: ss.R
'E:
': ss.B cSub, cMod, "pNmCmdBar,pLnBtn,pBtnSte", pNmCmdBar, pLnBtn, pBtnSte
'    Set_CmdBtnSte = True
'End Function
'Function Set_Colr_Chk(pChk As Access.CheckBox, pEnable As Boolean) As Boolean
'On Error Resume Next
'With pChk
'    If pEnable Then
'        .BorderColor = 65280
'    Else
'        .BorderColor = 13209
'    End If
'End With
'End Function
'Function Set_Colr_Lbl(pLbl As Label, pEnable As Boolean) As Boolean
'On Error Resume Next
'With pLbl
'    If pEnable Then
'        .BackColor = 65280
'        .ForeColor = 0
'    Else
'        .BackColor = 13209
'        .ForeColor = 16777215
'    End If
'End With
'End Function
'Function Set_CtlLayout(pCtl As Access.Control, Optional pLeft! = -1, Optional pTop! = -1, Optional pWdt! = -1, Optional pHgt! = -1) As Boolean
'Const cSub$ = "Set_CtlLayout"
'On Error GoTo R
'With pCtl
'    If pTop >= 0 Then .Top = pTop
'    If pLeft >= 0 Then .Left = pLeft
'    If pHgt >= 0 Then .Height = pHgt
'    If pWdt >= 0 Then .Width = pWdt
'End With
'Exit Function
'R: ss.R
'E:
': ss.B cSub, cMod, "pCtl,pLeft,pTop,pWdt,pHgt", ToStr_Ctl(pCtl), pLeft, pTop, pWdt, pHgt
'    Set_CtlLayout = True
'End Function
'Function Set_CtlPrp(pCtl As Access.Control, PrpNm$, pV) As Boolean
'Const cSub$ = "Set_CtlPrp"
'On Error GoTo R
'pCtl.Properties(PrpNm).Value = pV
'Exit Function
'R: ss.R
'E:
': ss.B cSub, cMod, "pCtl,PrpNm,pV", ToStr_Ctl(pCtl), PrpNm, pV
'    Set_CtlPrp = True
'End Function
'Function Set_CtlPrp_InFrm(pFrm As Access.Form, pTagSubStr$, PrpNm$, pV) As Boolean
'Const cSub$ = "Set_CtlPrp_InFrm"
'Dim iCtl As Access.Control: For Each iCtl In pFrm.Controls
'    If InStr(iCtl.Tag, pTagSubStr) > 0 Then Set_CtlPrp iCtl, PrpNm, pV
'Next
'End Function
'Function Set_CtlVisible(pFrm As Form, pVisibleTag$, Optional pInVisibleTag$) As Boolean
'Dim iCtl As Control
'On Error Resume Next
'For Each iCtl In pFrm.Controls
'    If InStr(iCtl.Tag, pVisibleTag$) Then iCtl.Visible = True
'    If InStr(iCtl.Tag, pInVisibleTag$) Then iCtl.Visible = False
'Next
'End Function
'Function Set_Cummulation(pRs As DAO.Recordset, pLoKey$, VFld$, pSetFld$) As Boolean
''Aim: Set Cummulation of <VFld> into <pSetFLd> with grouping as defined in list of key fields <pKeyFlds>
''Output: the field pRs->pSetFld will be Updated
''Input : pRs, pKeyFlds, VFld, pSetFld
'''pRs     : Assume it has been sorted in proper order
'''pKeyFlds: a list of key fields used as grouping the records in pRs (same records with pKeyFlds value considered as a group)
'''VFld : VFld is the value field name used to do the cummulation to set the pSetFld.  If ="", use 1 as value.
'''pSetFld : the field required to set
''Logic : For each group of records in pRs, the pSetFld will be set to cummulate the field VFld
''Example: in ATP.mdb: ATP_35_FullSetNew_3Upd_Qty_As_Cummulate_RunCode()
'''- Input table is : tmpATP_FullSetNew
'''                   FGDmdId / FG / CmpSupTypSeq / CmpSupTyp / DelveryDate / Cmp / Qty / RunningQty
'''- pRs      = currentdtable("tblATP_FullSetNew").openrecordset
'''             pRs.index = "PrimaryKey"
'''             pRs.PrimaryKey is : FGDmdId / FG / Cmp / CmpSupTypSeq / CmpSupTyp / DeliveryDate
'''- pKeyFlds = FGDmdId / FG / Cmp
'''- VFld  = "Qty"
'''- pSetFld  = RunningQty
'Dim mAnFldKey$(): mAnFldKey = Split(pLoKey, CtComma)
'Dim NKey%: NKey = Sz(mAnFldKey)
'ReDim mAyLasKeyVal(NKey - 1)
'Dim J As Byte: For J = 0 To NKey - 1
'    mAyLasKeyVal(J) = "xxxx"
'Next
'Dim mQ_Run As Double
'With pRs
'    While Not .EOF
'        If IsSamKey_ByAnFldKey(pRs, mAnFldKey, mAyLasKeyVal) Then
'            If VFld = "" Then
'                mQ_Run = mQ_Run + 1
'            Else
'                mQ_Run = mQ_Run + Nz(pRs.Fields(VFld).Value, 0)
'            End If
'        Else
'            If VFld = "" Then
'                mQ_Run = 1
'            Else
'                mQ_Run = Nz(pRs.Fields(VFld).Value, 0)
'            End If
'            For J = 0 To NKey - 1
'                mAyLasKeyVal(J) = pRs.Fields(mAnFldKey(J)).Value
'            Next
'        End If
'        .Edit
'        .Fields(pSetFld).Value = mQ_Run
'        .Update
'        .MoveNext
'    Wend
'End With
'End Function
'Function Set_Dbl0Dft(pNmt$) As Boolean
''Aim: set all double fields in {pNmt} to have 0 as default value
'Dim J%
'For J = 0 To CurrentDb.TableDefs(pNmt).Fields.Count - 1
'    If CurrentDb.TableDefs(pNmt).Fields(J).Type = dbDouble Then Set_FldDftV pNmt, CurrentDb.TableDefs(pNmt).Fields(J).Name, 0
'Next
'End Function
'Function Set_DocPrp(pWb As Workbook, pDocPrp As tDocPrp) As Boolean
'With pDocPrp
'    pWb.BuiltinDocumentProperties("Title").Value = .NmRpt
'    pWb.BuiltinDocumentProperties("Subject").Value = .NmRptSht & "-" & .NmSess
'    pWb.BuiltinDocumentProperties("Author").Value = "Johnson Cheung"
'    pWb.BuiltinDocumentProperties("Comments").Value = _
'        "Generated @ " & Format(Now(), "yyyy/mm/dd hh:nn:ss") & vbLf & _
'        "Generated by " & CurrentDb.Name & vbLf & _
'        "Data name: " & .NmData & vbLf & _
'        "ExtraPrm : " & .ExtraPrm
'    pWb.BuiltinDocumentProperties("Keywords").Value = .NmRptSht & CtComma & .NmSess
'End With
'End Function
'
'Function Set_FfnPDF(pFfnPDF$) As Boolean
'Dim mDir$: mDir = Fct.Nam_DirNam(pFfnPDF)
'Dim mFn$:  mFn = Fct.Nam_FilNam(pFfnPDF)
'With gPDF
'    .cOption("UseAutosave") = 1
'    .cOption("UseAutosaveDirectory") = 1
'    .cOption("AutosaveDirectory") = mDir
'    .cOption("AutosaveFilename") = mFn
'    .cOption("AutosaveFormat") = 0                            ' 0 = PDF
'    .cStart
'End With
'End Function
'
'Function Set_FilRO(pFfn$) As Boolean
'Const cSub$ = "Set_FilRO"
'On Error GoTo R
'FileSystem.SetAttr pFfn, vbReadOnly
'Exit Function
'R: ss.R
'E:
': ss.B cSub, cMod, "pFfn", pFfn
'    Set_FilRO = True
'End Function
'Function Set_FilRW(pFfn$) As Boolean
'Const cSub$ = "Set_FilRW"
'On Error GoTo R
'FileSystem.SetAttr pFfn, vbNormal
'Exit Function
'R: ss.R
'E:
': ss.B cSub, cMod, "pFfn", pFfn
'    Set_FilRW = True
'End Function
'Function Set_FldDftV(pNmt$, pFldNm$, pDftV) As Boolean
'Const cSub$ = "Set_FldDftV"
'On Error GoTo R
'CurrentDb.TableDefs(pNmt).Fields(pFldNm).DefaultValue = pDftV
'Exit Function
'R: ss.R
'E: Set_FldDftV = True: ss.B cSub, cMod, "pNmt,pFldNm,pDftV", pNmt, pFldNm, pDftV
'End Function
'Function Set_Formula(Rg As Range, pNRow&, pFormula$) As Boolean
''Aim: Copy formula at {Rg} download {pNRow} (including the row of {Rg}
'Const cSub$ = "Set_Formula"
'If pNRow <= 0 Then Exit Function
'On Error GoTo R
'With Rg(1, 1)
'    .Formula = pFormula
'    .Copy
'End With
'Dim mWs As Worksheet: Set mWs = Rg.Parent
'mWs.Range(Rg(2, 1), Rg(pNRow, 1)).PasteSpecial xlPasteFormulas
'Exit Function
'R: ss.R
'E: Set_Formula = True: ss.B cSub, cMod, "Rg,NRow,pFormula", ToStr_Rge(Rg), pNRow, pFormula
'End Function
'Function Set_Formula_SumNxtN(pWs As Worksheet, pCno As Byte, pRnoBeg&, pNRow&, pNCol As Byte) As Boolean
'Const cSub$ = "Set_Formula_SumNxtN"
'Dim mNxt1$: mNxt1 = WsCnoCol(pCno + 1) & pRnoBeg
'Dim mNxtN$: mNxtN = WsCnoCol(pCno + pNCol) & pRnoBeg
'Dim mCol$: mCol = WsCnoCol(pCno)
'With pWs.Range(mCol & pRnoBeg)
'    .Formula = Fmt_Str("=Sum({0}:{1})", mNxt1, mNxtN)
'    .Copy
'End With
'pWs.Range(mCol & pRnoBeg & ":" & mCol & pRnoBeg + pNRow - 1).PasteSpecial xlPasteFormulas
'Exit Function
'R: ss.R
'E: Set_Formula_SumNxtN = True: ss.B cSub, cMod, "pWs,pCno,pRnoBeg,pNRow,pNCol", ToStr_Ws(pWs), pCno, pRnoBeg, pNRow, pNCol
'End Function
'Sub WsSetFze(A As Worksheet, Adr$)
'With A.Range(Adr)
'    .Activate
'    .Select
'End With
'ActiveWindow.FreezePanes = True
'End Sub
'Sub RgSetHypLnk(A As Excel.Range)
''Aim: Set any cells within the {Rg} to hyper link to A1 of worksheet if they have the same value
'Const cSub$ = "Set_HypLnk"
'Dim mWs As Worksheet: Set mWs = Rg.Worksheet
'Dim mWb As Workbook: Set mWb = mWs.Parent
'Dim mAnWs$(): If Fnd_AnWs_ByWb(mAnWs, mWb) Then GoTo E
'Dim N%: N = Sz(mAnWs)
'Dim iCell As Range, V, J%
'For Each iCell In Rg
'    V = iCell.Value
'    If VarType(V) = vbString Then
'        V = Left(V, 31)
'        For J = 0 To N - 1
'            If V = mAnWs(J) Then Call mWs.Hyperlinks.Add(iCell, "", CtSngQ & mWb.Sheets(mAnWs(J)).Name & "'!A1")
'        Next
'    End If
'Next
'Exit Sub
'E:
': ss.B cSub, cMod, "pAy", ToStr_Rge(Rg)
'    Set_HypLnk = True
'End Sub
'Function Set_HypLnk__Tst()
'Const cFfn$ = "c:\temp\a.xls"
'Dim mWb As Workbook: Set mWb = G.gXls.Workbooks.Open(cFfn)
'Dim mWs As Worksheet: Set mWs = mWb.Sheets("Index")
'mWb.Application.Visible = True
'If Set_HypLnk(mWs.Range("A1:E200")) Then Stop
'End Function
'Function Set_Lck(pFrm As Access.Form, pLck As Boolean, Optional pAlwAdd As Boolean = False, Optional pAlwDlt As Boolean = False) As Boolean
''Aim: Set all controls in {pFrm} as lock
'Const cSub$ = "Set_Lck"
'Dim iCtl As Access.Control: For Each iCtl In pFrm.Controls
'    If Not Visible Then GoTo Nxt
'    Dim mLck As Boolean: If iCtl.Tag = "Edt" Then mLck = pLck Else mLck = True
'    Select Case TypeName(iCtl)
'    Case "Label"
'        If IsEnd(iCtl.Name, "_Lbl") Then GoTo Nxt
'        If Set_LckLbl(iCtl, mLck) Then ss.A 1: GoTo E
'    Case "TextBox":  If Set_LckTBox(iCtl, mLck) Then ss.A 2: GoTo E
'    Case "Check":    If Set_LckChkB(iCtl, mLck) Then ss.A 3: GoTo E
'    Case "ComboBox": If Set_LckCBox(iCtl, mLck) Then ss.A 4: GoTo E
'    End Select
'Nxt:
'Next
'With pFrm
'    If pLck Then
'        .AllowEdits = False
'        .AllowAdditions = False
'        .AllowDeletions = False
'    Else
'        .AllowEdits = True
'        .AllowAdditions = pAlwAdd
'        .AllowDeletions = pAlwDlt
'    End If
'End With
'pFrm.Repaint
'Exit Function
'R: ss.R
'E: Set_Lck = True: ss.B cSub, cMod, "pFrm,pLck,pAlwAdd,pAlwDlt", ToStr_Frm(pFrm), pLck, pAlwAdd, pAlwDlt
'End Function
'
'Function Set_Lck__Tst()
'Const cNmFrm$ = "frmIIC_Tst"
'Dim mFrm As Access.Form: If FrmOpn(cNmFrm, , , mFrm) Then Stop: GoTo E
'If Set_Lck(mFrm, False) Then Stop: GoTo E
'Stop
'If Set_Lck(mFrm, True) Then Stop: GoTo E
'Stop
'If Set_Lck(mFrm, False) Then Stop: GoTo E
'Stop
'GoTo X
'E: Set_Lck_Tst = True
'X: Cls_Frm cNmFrm
'End Function
'
'Function Set_LckCBox(pCBox As Access.ComboBox, pLck As Boolean) As Boolean
'Const cSub$ = "Set_LckCBox"
'pCBox.Locked = pLck
'pCBox.ForeColor = 0
'pCBox.TabStop = Not pLck
''pCBox.ForeColor = IIf(pLck, 255, 0)
'End Function
'Function Set_LckChkB(pChkB As Access.CheckBox, pLck As Boolean) As Boolean
'Const cSub$ = "Set_LckChkB"
'pChkB.Locked = pLck
'pChkB.BorderColor = IIf(pLck, 13209, 65280)
'pChkB.TabStop = Not pLck
'End Function
'Function Set_LckLbl(pLbl As Access.Label, pLck As Boolean) As Boolean
'On Error Resume Next
'pLbl.ForeColor = IIf(pLck, 16777215, 0)
'pLbl.BackColor = IIf(pLck, 13209, 65280)
'End Function
'
'Function Set_LckLbl__Tst()
'Const cNmFrm$ = "frmIIC_Tst"
'Dim mFrm As Access.Form: If FrmOpn(cNmFrm, , , mFrm) Then Stop: GoTo E
'Dim mLbl As Access.Label: Set mLbl = mFrm.Controls("ICGL_Label")
'If Set_LckLbl(mLbl, False) Then Stop: GoTo E
'Exit Function
'E: Set_LckLbl_Tst = True
'End Function
'
'Function Set_LckTBox(pTBox As Access.TextBox, pLck As Boolean) As Boolean
'Const cSub$ = "Set_LckTBox"
'pTBox.Locked = pLck
'pTBox.ForeColor = 0
'pTBox.TabStop = Not pLck
'Dim mLbl As Access.Label: If Fnd_Lbl(mLbl, pTBox) Then Exit Function
'Set_LckLbl mLbl, pLck
'End Function
'
'Function Set_LckTBox__Tst()
'Const cNmFrm$ = "frmIIC_Tst"
'Dim mFrm As Access.Form: If FrmOpn(cNmFrm, , , mFrm) Then Stop: GoTo E
'Dim mTBox As Access.TextBox: Set mTBox = mFrm.Controls("ICGL")
'If Set_LckTBox(mTBox, False) Then Stop: GoTo E
'Exit Function
'E: Set_LckTBox_Tst = True
'End Function
'
'Function Set_LstCtlLayout(pFrm As Access.Form, pLnCtl$, Optional pLeft! = -1, Optional pTop! = -1, Optional pWdt! = -1, Optional pHgt! = -1) As Boolean
'Const cSub$ = "Set_LstCtlLayout"
'Dim mAnCtl$(): mAnCtl = Split(pLnCtl, CtComma)
'Dim J%: For J = 0 To Sz(mAnCtl) - 1
'    Dim iCtl As Access.Control: If Fnd_Ctl(iCtl, pFrm, mAnCtl(J)) Then GoTo Nxt
'    Set_CtlLayout iCtl, pLeft, pTop, pWdt, pHgt
'Nxt:
'Next
'End Function
'Function Set_Nm_InWb(pWb As Workbook, pNm$, pReferTo$) As Boolean
'Const cSub$ = "Set_Set_Nm_InWb"
'On Error GoTo R
'Dim mNm As Name: If IsWbNm(pWb, pNm, mNm) Then mNm.RefersTo = pReferTo: Exit Function
'pWb.Names.Add pNm, pReferTo$
'GoTo X
'R: ss.R
'E: Set_Nm_InWb = True: ss.B cSub, cMod, "pWb,pNm,pReferTo", WbToStr(pWb), pNm, pReferTo
'X:
'End Function
'Function Set_Nm_InWs(pWs As Worksheet, pNm$, pReferTo$) As Boolean
'Const cSub$ = "Set_Set_Nm_InWs"
'On Error GoTo R
'Dim mNm As Name: If IsWsNm(pWs, pNm, mNm) Then mNm.RefersTo = pReferTo: Exit Function
'pWs.Names.Add pNm, pReferTo$
'GoTo X
'R: ss.R
'E: Set_Nm_InWs = True: ss.B cSub, cMod, "pWs,pNm,pReferTo", ToStr_Ws(pWs), pNm, pReferTo
'X:
'End Function
'Function Set_PdfPrt(pSetPdfPrt As Boolean) As Boolean
'Const cSub$ = "Set_PdfPrt"
'On Error GoTo R
'Static xSavPrt$
'With gWrd
'    If pSetPdfPrt Then
'        If Left(.ActivePrinter, 10) = "PDFCreator" Then Exit Function
'        xSavPrt = .ActivePrinter
'        .ActivePrinter = "PDFCreator"
'        Exit Function
'    End If
'    If xSavPrt <> "" Then If .ActivePrinter <> xSavPrt Then .ActivePrinter = xSavPrt
'    xSavPrt = ""
'    Exit Function
'End With
'Exit Function
'R: ss.R
'
'E:
': ss.B cSub, cMod, "pSetPdfPrt", pSetPdfPrt
'    Set_PdfPrt = True
'End Function
'
'Function Set_PdfPrt__Tst()
'Const cSub$ = "Set_PdfPrt_Tst"
'Shw_Dbg cSub, cMod
'Dim J%: For J = 0 To 10
'    Debug.Print J
'    Set_PdfPrt True
'    Set_PdfPrt False
'Next
'End Function
'
'Function Set_Pf_OfWb(pWb As Workbook, pLnPf$, Optional pOrientation As XlOrientation = xlHidden) As Boolean
''Aim: Hide the pivot fields {pLnPf} inside the {pWb}
''Param: pLnPf is a list Pivot Field Name separated by CtComma
'Dim J As Byte, AnPf$(): AnPf = Split(pLnPf, CtComma)
'Dim iWs As Worksheet
'For Each iWs In pWb.Worksheets
'    Dim iPt As PivotTable
'    For Each iPt In iWs.PivotTables
'        For J = LBound(AnPf) To UBound(AnPf)
'            On Error Resume Next
'            iPt.PivotFields(AnPf(J)).Orientation = pOrientation
'            On Error GoTo 0
'        Next
'    Next
'Next
'End Function
'
'
'Function QrySetPrp(pQry As QueryDef, PrpNm$, V$) As Boolean
'If PrpNm = "Description" And V = "" Then
'    On Error Resume Next
'    pQry.Properties.Delete PrpNm: Exit Function
'    Exit Function
'End If
'On Error GoTo R
'
'pQry.Properties(PrpNm).Value = V
'Exit Function
'R: ss.R
'    Dim mPrp As DAO.Property: Set mPrp = pQry.CreateProperty(PrpNm, DAO.DataTypeEnum.dbText, V)
'    pQry.Properties.Append mPrp
'End Function
'Function QrySetPrp_Bool(pQry As QueryDef, PrpNm$, V As Boolean) As Boolean
'On Error GoTo R
'pQry.Properties(PrpNm).Value = V
'Exit Function
'R: ss.R
'    Dim mPrp As DAO.Property: Set mPrp = pQry.CreateProperty(PrpNm, DAO.DataTypeEnum.dbBoolean, V)
'    pQry.Properties.Append mPrp
'End Function
'Function Set_Sno(pNmt$, Optional pNmFldSno$ = "Sno", Optional pOrdBy$ = "") As Boolean
'Const cSub$ = "Set_Sno"
''-- Fill in <<pNmSeqFld>> starting from 1 by using PrimaryKey as the key
'On Error GoTo R
'Dim mNmt$: mNmt = Q_S(pNmt, "[]")
'Dim mRs As DAO.Recordset: If Opn_Rs(mRs, Fmt_Str("Select {0} from {1}{2}", pNmFldSno, mNmt, SqsOrdBy(pOrdBy))) Then ss.A 1: GoTo E
'Set_Sno = Set_Sno_ByRs(mRs, pNmFldSno$)
'Exit Function
'R: ss.R
'E: Set_Sno = True
': ss.B cSub, cMod, "pNmt,pNmFldSno$", pNmt, pNmFldSno
'End Function
'
'Function Set_Sno__Tst()
'TblCrt_ByFldDclStr "#Tmp", "aa Text 10, Sno Long") Then Stop: GoTo E
'Dim J%
'For J = 0 To 10
'    If Run_Sql("Insert into [#Tmp] (aa) values ('{0}')") Then Stop: GoTo E
'Next
'If Set_Sno("#Tmp") Then Stop: GoTo E
'Exit Function
'E:
'    Set_Sno_Tst = True
'End Function
'
'Function Set_Sno_ByRs(pRs As DAO.Recordset, pNmFldSno$) As Boolean
'Const cSub$ = "Set_Sno_ByRs"
'On Error GoTo R
'With pRs
'    Dim mSno&
'    While Not .EOF
'        .Edit
'        mSno = mSno + 1
'        .Fields(pNmFldSno$).Value = mSno
'        .Update
'        .MoveNext
'    Wend
'    .Close
'End With
'Exit Function
'R: ss.R
'E: Set_Sno_ByRs = True
': ss.B cSub, cMod, "pNmFldSno$", pNmFldSno$
'End Function
'Function Set_AyKv_ByRs(oAyKv(), pRs As DAO.Recordset) As Boolean
''Aim: Set first N fields value of {pRs} to {oAyKv}
'Const cSub$ = "Set_AyKv_ByRs"
'On Error GoTo R
'Dim J%, mN%: mN = Sz(oAyKv)
'For J = 0 To mN - 1
'    oAyKv(J) = pRs.Fields(J).Value
'Next
'Exit Function
'R: ss.R
'E: Set_AyKv_ByRs = True: ss.B cSub, cMod, "Siz(oAyKv),pRs", mN, ToStr_Rs_NmFld(pRs)
'End Function
'Function Set_Sno_wGp(pNmt$, FnStrGp$, Optional pOrdBy$ = "", Optional pNmFldSeq$ = "Sno") As Boolean
'Const cSub$ = "Set_Sno_wGp"
'On Error GoTo R
'Dim mNmt$: mNmt = Q_SqBkt(pNmt)
'Dim mSql$: mSql = Fmt_Str("Select {2},{0} from {1} Order by {2}{3}", pNmFldSeq, mNmt, FnStrGp, Cv_Str(pOrdBy, ","))
'Dim mRs As DAO.Recordset: If Opn_Rs(mRs, mSql) Then ss.A 1: GoTo E
'Dim mAnFldGp$(): mAnFldGp = Split(FnStrGp, ",")
'Dim NGp%: NGp = Sz(mAnFldGp)
'ReDim mAyKvLas(NGp - 1)
'Dim mSno%: mSno = 0
'With mRs
'    While Not .EOF
'        If Not IsSamKey(mRs, mAyKvLas) Then
'            If Set_AyKv_ByRs(mAyKvLas, mRs) Then ss.A 2: GoTo E
'            mSno = 0
'        End If
'        mSno = mSno + 10
'        .Edit
'        .Fields(pNmFldSeq).Value = mSno
'        .Update
'        .MoveNext
'    Wend
'End With
'GoTo X
'R: ss.R
'E: Set_Sno_wGp = True: ss.B cSub, cMod, "pNmt,FnStrGp,pOrdBy,pNmFldSeq", pNmt, FnStrGp, pOrdBy, pNmFldSeq
'X:
'    RsCls mRs
'End Function
'Function Set_SetCmdBarEnable(pNmCmdBar$, pEnable As Boolean) As Boolean
'Dim iCtl As CtCommandBarControl
'On Error Resume Next
'For Each iCtl In Application.CtCommandBars(pNmCmdBar).Controls
'    iCtl.Enabled = pEnable
'Next
'On Error GoTo 0
'End Function
'Function Set_SetColWidth(pWs As Worksheet, pColWidthLst$) As Boolean
'Dim Ay$(), J%, iColNo As Byte, mRange As Range
'Ay = Split(pColWidthLst, CtComma)
'With pWs
'    For J = LBound(Ay) To UBound(Ay)
'        iColNo = iColNo + 1
'        Set mRange = .Cells(1, iColNo)
'        mRange.EntireColumn.ColumnWidth = Ay(J)
'    Next
'End With
'End Function
'Function Set_SubTot(pWs As Worksheet, pCno As Byte, pRnoBeg&, pNRow&, Optional pFctNo As Byte = 9) As Boolean
'Const cSub$ = "Set_SubTot"
'On Error GoTo R
'Dim mCol$: mCol = WsCnoCol(pCno)
'pWs.Range(mCol & (pRnoBeg + pNRow)).Formula = Fmt_Str("=SUBTOTAL({0},{1})", pFctNo, mCol & pRnoBeg & ":" & mCol & pRnoBeg + pNRow - 1)
'Exit Function
'R: ss.R
'E: Set_SubTot = True: ss.B cSub, cMod, "pWs,pCno,pRnoBeg,pNRow,pFctNo", ToStr_Ws(pWs), pCno, pRnoBeg, pNRow, pFctNo
'End Function
'Function Set_TblSeqInDesc(pQryPfx$) As Boolean
''Aim   : Set first 2 char of desc of the table tmpQQQ_TTT to NN
''Assume: - The make table query qryQQQ_NN_1_Fm_XXXX will generate tmp table tmpQQQ_TTT.
''        - The select query is  qryQQQ_NN_0_TTT.
''
''Logic : For each "Select query" in format of qryQQQ_NN_0_TTT
''          If tmpQQQ_TTT exist, Find nn & set the desc
''        Next
'Const cSub$ = "Set_TblSeqInDesc"
'Dim L As Byte: L = Len(pQryPfx): If L = 0 Then ss.A 1, "pQryPfx cannot zero length": GoTo E
'Dim iQry As QueryDef, iNmq$
'For Each iQry In CurrentDb.QueryDefs
'    If Left(iQry.Name, L) <> pQryPfx Then GoTo Nxt
'    If iQry.Type <> DAO.QueryDefTypeEnum.dbQSelect Then GoTo Nxt
'    iNmq = iQry.Name
'    'Get iNN as II_0_PPPP
'    'Get iStep as 0
'    Dim iNN$:      iNN = Mid$(iNmq, L + 2)
'    Dim iStep$:    iStep = Mid$(iNN, 4, 1)
'    If iStep <> "0" Then GoTo Nxt
'    'Get iII as II
'    'Get iPPPP as PPPP
'    Dim iII$:      iII = Left(iNN, 2)
'    Dim iPPPP$:    iPPPP = Mid$(iNN, 6)
'    '-- Get iTmpNmt as tmpMMMM_PPPP
'    Dim iTmpNmt$:  iTmpNmt = "tmp" & Mid$(pQryPfx, 4) & "_" & iPPPP
'    If IsTbl(iTmpNmt) Then
'        Call Set_TblSeqInDesc_SetDesc(iTmpNmt, iII)
'        Debug.Print iNmq; " "; iTmpNmt; " is set to -----> " & iII
'    Else
'        Debug.Print iNmq; " "; iTmpNmt; " does not exist"
'    End If
'Nxt:
'Next
'Exit Function
'E: Set_TblSeqInDesc = True: ss.B cSub, cMod, "pQryPfx", pQryPfx
'End Function
'Function Set_TblSeqInDesc_SetDesc(pNmt$, pNN$) As Boolean
'Dim mDesc$: mDesc = Fnd_Prp(pNmt, acTable, "Description")
'Set_Prp pNmt, acTable, "Description", pNN & Mid$(mDesc, 3)
'End Function
'Function Set_TblZero2Null(pNmt, FnStr_SubStr) As Boolean
'Dim mAnFld_SubStr$(): mAnFld_SubStr = Split(FnStr_SubStr, CtComma)
'With CurrentDb.TableDefs(pNmt).OpenRecordset
'    While Not .EOF
'        Dim iSubStr As Byte: For iSubStr = LBound(mAnFld_SubStr) To UBound(mAnFld_SubStr)
'            Dim iFld As DAO.Field: For Each iFld In .Fields
'                .Edit
'                If InStr(iFld.Name, mAnFld_SubStr(iSubStr)) > 0 Then
'                    If iFld.Value = 0 Then iFld.Value = Null
'                End If
'                .Update
'            Next
'        Next
'        .MoveNext
'    Wend
'    .Close
'End With
'End Function
'Function Set_Lv2ColAtEnd(oRgeLv As Range, pLv$, pWs As Worksheet _
'    , Optional pRow1Val$ = "WsOfFollowRge" _
'    ) As Boolean
''Aim:
'
'Const cSub$ = "Set_Lv2ColAtEnd"
'On Error GoTo R
''Do Set Row1 & pLv in an empty column & Set oRgeLv
'Do
'    'Do Find mCno_Empty
'    Dim mCno_Empty As Byte
'    Do
'        mCno_Empty = Fnd_Cno_EmptyCell_InRow(pWs, , 255, 1)
'        If mCno_Empty = 0 Then ss.A 2: GoTo E
'    Loop Until True
'    With pWs
'    'Set Row1
'        pWs.Cells(1, mCno_Empty).Value = pRow1Val
'        'Set Lv to empty column
'        Dim mAy$(): mAy = Split(pLv, CtComma)
'        Dim J%, N%: N = Sz(mAy)
'        For J = 0 To N - 1
'            .Cells(2 + J, mCno_Empty).Value = mAy(J)
'        Next
'        'Set oRgeLv
'        Set oRgeLv = .Range(.Cells(2, mCno_Empty), .Cells(J + N - 1, mCno_Empty))
'    End With
'Loop Until True
'Exit Function
'R: ss.R
'E: Set_Lv2ColAtEnd = True: ss.B cSub, cMod, "pLv,pWs,pRow1Val", pLv, ToStr_Ws(pWs), pRow1Val
'End Function
'Function Set_RgeVdt_ByLv(Rg As Range, pLv$ _
'    , Optional pInputTit$ = "Enter value or leave blank" _
'    , Optional pInputMsg$ = "Enter one of the value in the list or leave it blank." _
'    , Optional pErrTit$ = "Not in the List" _
'    , Optional pErrMsg = "Please enter a value in list or leave it blank" _
'    ) As Boolean
''Aim: Set the validation of {Rg} to select a list of value {pLv}.
''     'The list of value' will be the stored in the avaliable column of ws [SelectionList]
''          Ws [SelectionList] Row1=Ws Name, Row2=Rge that will use to the list to select value, Row3 and onward will be the selection value
'Const cSub$ = "Set_RgeVdt_ByLv"
'
'' Do Build mRgeLv: 'The list of value'
'Dim mRgeLv As Range: If Set_Lv2ColAtEnd(mRgeLv, pLv, Rg.Worksheet, Rg.Address) Then ss.A 1: GoTo E
'
'' Do Set Vdt of Rg
'Do
'    On Error GoTo R
'    With Rg.Validation
'        .Delete
'        Dim mFormula$: mFormula = "=" & mRgeLv.Address
'        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=mFormula
'        .IgnoreBlank = True
'        .InCellDropdown = True
'        .InputTitle = pInputTit
'        On Error Resume Next
'        .InputMessage = pInputMsg
'        On Error GoTo R
'        .ErrorTitle = pErrTit
'        .ErrorMessage = pErrMsg
'        .ShowInput = True
'        .ShowError = True
'    End With
'Loop Until True
'Exit Function
'R: ss.R
'E: Set_RgeVdt_ByLv = True: ss.B cSub, cMod, "Rg,pLv", ToStr_Rge(Rg), pInputTit, pInputMsg, pErrTit, pErrMsg, pLv, pInputTit, pInputMsg, pErrTit, pErrMsg
'End Function
'
'Function Set_RgeVdt_ByLv__Tst()
'Const cSub$ = "Set_RgeVdt_ByLv_Tst"
'Dim mRge As Range, mLv$
'Dim mRslt As Boolean, mCase As Byte: mCase = 1
'
'Dim mWb As Workbook: If Crt_Wb(mWb, "c:\aa.xls", True) Then ss.A 1: GoTo E
'mWb.Application.Visible = True
'Select Case mCase
'Case 1
'    Set mRge = mWb.Sheets(1).Range("A1:D5")
'    mLv = "aa,bb,cc,11,22,33"
'End Select
'mRslt = Set_RgeVdt_ByLv(mRge, mLv)
'Shw_Dbg cSub, cMod, "mRslt, mRge, mLv", mRslt, ToStr_Rge(mRge), mLv
'Exit Function
'R: ss.R
'E: Set_RgeVdt_ByLv_Tst = True: ss.B cSub, cMod
'End Function
'
'
'
