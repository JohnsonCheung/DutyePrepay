Attribute VB_Name = "ZZ_xGen"

'Option Compare Text
'Option Explicit
'Option Base 0
'Const cMod$ = cLib & ".xGen"
'Function Gen_jj_Xla() As Boolean
'Const cSub$ = "Gen_jj_Xla"
'If Exp_Prj(, , , True) Then ss.A 1: GoTo E
'If Gen_PgmXls_FmDir Then ss.A 2: GoTo E
'If Cpy_Fil(Sdir_Doc & "\Pgm\xla", Sdir_Hom & "AppDef_Meta\jjNew.xla", True) Then ss.A 3: GoTo E
'Exit Function
'E: Gen_jj_Xla = True: ss.B cSub, cMod
'End Function
'Function Gen_PgmXls_FmDir(Optional pDir$ = "") As Boolean
''Aim: generate Xls file in same dir as {mDir}.  Assuming the {mDir} contains files of modules (xx.xxx.bas) and classes (xx.xxx.cls)
''     xx is project name.  xxx is module/class name
'Const cSub$ = "Gen_PgmXls_FmDir"
'On Error GoTo R
'Dim mDir$: If mDir = "" Then mDir = Sdir_ExpPgm Else mDir = pDir
'If Dlt_Fil_BySfx(mDir, ".Xla") Then ss.A 1: GoTo E
'Dim mAyFn$(): If Fnd_AyFn(mAyFn, mDir, "*.bas,*.cls,*.Reference.Txt") Then ss.A 2: GoTo E
'Dim mNmPrjLas$
'Dim J%, mA$
'For J = 0 To Sz(mAyFn) - 1
'    StsShw "Building module " & mAyFn(J) & " ..."
'    Dim mNmPrj$, mNmm$, mExt$
'    If Brk_Str_To3Seg(mNmPrj, mNmm, mExt, mAyFn(J), ".") Then ss.A 3: GoTo E
'
'    If mNmPrj <> mNmPrjLas Then
'        Dim mFx$: mFx = mDir & mNmPrj & ".Xla"
'        Dim mWb As Workbook: If Not IsNothing(mWb) Then If Sav_Wb_AsXla(mWb) Then ss.A 4: GoTo E
'        If Crt_Wb(mWb, mFx, True, "Sheet1") Then ss.A 5: GoTo E
'        Dim mPrj As vbproject: Set mPrj = mWb.vbproject
'        mPrj.Name = mNmPrj
'        mNmPrjLas = mNmPrj
'    End If
'    Dim mFfn$: mFfn = mDir & mAyFn(J)
'    If mNmm = "Reference" And mExt = "Txt" Then
'        If Add_Rf(mPrj, mFfn) Then ss.A 7: GoTo E
'    Else
'        If Add_Md_ToPrj(mPrj, mFfn) Then mA = Add_Str(mA, mAyFn(J))
'    End If
'Next
'If Not IsNothing(mWb) Then If Sav_Wb_AsXla(mWb) Then ss.A 1: GoTo E
'If mA <> "" Then ss.A 9, "Some file cannot be added as module/class/form/report", , "The Fn", mA: GoTo E
'GoTo X
'R: ss.R
'E: Gen_PgmXls_FmDir = True: ss.B cSub, cMod, "pDir,mDir", pDir, mDir
'X: Clr_Sts
'   Cls_Wb mWb, False, True
'End Function

'Function Gen_PgmXls_FmDir__Tst()
'Const cSub$ = "Gen_PgmXls_FmDir_Tst"
'If Gen_PgmXls_FmDir("P:\Documents\Pgm\") Then ss.A 1: GoTo E
'Exit Function
'E: Gen_PgmXls_FmDir_Tst = True: ss.B cSub, cMod
'End Function

'Function Gen_PgmAcs_FmDir(pDir$) As Boolean
''Aim: generate Acs file in same dir as {pDir}.  Assuming the {pDir} contains files of modules (xx.xxx.bas) and classes (xx.xxx.cls)
''     xx is project name.  xxx is module/class name
'Const cSub$ = "Gen_PgmAcs_FmDir"
'On Error GoTo R
'If Dlt_Fil_BySfx(pDir, ".mdb") Then ss.A 1: GoTo E
'Dim mAyFn$(): If Fnd_AyFn(mAyFn, pDir, "*.*") Then ss.A 2: GoTo E
'Dim mNmPrjLas$
'Dim J%, mA$
'Dim mAcs As Access.Application: Set mAcs = G.gAcs
'For J = 0 To Sz(mAyFn) - 1
'    StsShw "Building module " & mAyFn(J) & " ..."
'    Dim mNmPrj$, mNmm$, mExt$
'    If Brk_Str_To3Seg(mNmPrj, mNmm, mExt, mAyFn(J), ".") Then ss.A 3: GoTo E
'
'    If mNmPrj <> mNmPrjLas Then
'        Set_Silent: Compile_Acs mAcs: Set_Silent_Rst
'        If Cls_CurDb(mAcs) Then ss.A 4: GoTo E
'        Dim mFb$: mFb = pDir & mNmPrj & ".Mdb": FbNew mFb) Then ss.A 5: GoTo E
'        If Opn_CurDb(mAcs, mFb) Then ss.A 6: GoTo E
'        Dim mPrj As vbproject: Set mPrj = mAcs.Vbe.ActiveVBProject
'        mPrj.Name = mNmPrj
'        mNmPrjLas = mNmPrj
'    End If
'    Dim mFfn$: mFfn = pDir & mAyFn(J)
'    If mNmm = "Reference" And mExt = "Txt" Then
'        If Add_Rf(mPrj, mFfn) Then ss.A 7: GoTo E
'    Else
'        If Add_Md(mAcs, mFfn) Then mA = Add_Str(mA, mAyFn(J))
'    End If
'Next
'Set_Silent: Compile_Acs mAcs: Set_Silent_Rst
'If Cls_CurDb(mAcs) Then ss.A 8: GoTo E
'If mA <> "" Then ss.A 9, "Some file cannot be added as module/class/form/report", , "The Fn", mA: GoTo E
'GoTo X
'R: ss.R
'E: Gen_PgmAcs_FmDir = True: ss.B cSub, cMod, "pDir", pDir
'X:
'    Cls_CurDb mAcs
'    Clr_Sts
'End Function

'Function Gen_PgmAcs_FmDir__Tst()
'Const cSub$ = "Gen_PgmAcs_FmDir_Tst"
'If Gen_PgmAcs_FmDir("P:\Documents\Pgm\") Then ss.A 1: GoTo E
'Exit Function
'E: Gen_PgmAcs_FmDir_Tst = True: ss.B cSub, cMod
'End Function

'Function Gen_Doc(Optional pLikNmm$ = "*", Optional pLikNmq$ = "", Optional pLikFrm$ = "", _
'    Optional pLikNmPrc$ = "*", Optional pLikFrmProc$ = "*") As Boolean
''Aim: Show all module's documentation into an Xls name <MdbNam>_Doc.xls in same directory
'Const cSub$ = "Gen_Doc"
'On Error GoTo R
''== Start
'If pLikFrm = "" And pLikNmm = "" And pLikNmq = "" Then ss.A 1, "At least a prefix (Frm/Mod/Qry) must be given": GoTo E
'
''Prepare *_#.csv (*=MdbNam; #=Qry,Mod,Frm)
'''Kill *_Doc.xls, *_#.csv
'Dim mFxDocPfx$: mFxDocPfx = Sdir_Doc & Sffn_Cur
'
'''Open *_Doc.csv & *_Qry_Doc.csv for output
'Dim mF As Byte, mAnm$(), J%, mMd As CodeModule
'''GenDoc for Modules into text file {<MdbNam>_Mod.csv}
'Dim mPrj As vbproject: Set mPrj = Application.Vbe.ActiveVBProject
'If pLikNmm <> "" Then
'    '''Open file {<MdbNam>_Doc.csv}
'    '''Write First Line :"#",<MdbNam_Full>,"Modules"
'    '''Write Header Line: "#", "Module", "Proc", "Line", "Remark"
'    If Opn_Fil_ForOutput(mF, mFxDocPfx & "_Mod.csv", True) Then ss.A 1: GoTo E
'    Write #mF, "#", CurrentDb.Name, "Modules"
'    Write #mF, "#", "Module", "Proc", "Line", "Remark"
'
'    '''Put all Modules of given prefix {p.PfxMod} in a Collection {mColl} and sort it
'    If Fnd_Anm_ByPrj(mAnm, mPrj, True) Then ss.A 2: GoTo E
'    For J = 0 To Sz(mAnm) - 1
'        If Fnd_Md(mMd, mPrj, mAnm(J)) Then ss.A 3: GoTo E
'        If Gen_Doc_For1Mod(mF, mMd, pLikNmPrc) Then ss.A 4: GoTo E
'    Next
'    Close #mF
'End If
'
'''GenDoc for Forms into text file {<MdbNam>_Frm.csv}
'If pLikFrm <> "" Then
'    '''Open file {<MdbNam>_Doc.csv}
'    '''Write First Line :"#",<MdbNam_Full>,"Forms"
'    '''Write Header Line: "#", "Form", "Proc", "Line", "Remark"
'    If Opn_Fil_ForOutput(mF, mFxDocPfx & "_Frm.csv", True) Then ss.A 2: GoTo E
'    Write #mF, "#", CurrentDb.Name, "Forms"
'    Write #mF, "#", "Form", "Proc", "Line", "Remark"
'
'    '''Put all Forms of given prefix {p.PfxFrm} in a Collection {mColl} and sort it
'    If Fnd_Anm_ByPrj(mAnm, mPrj, , True) Then ss.A 2: GoTo E
'    For J = 0 To Sz(mAnm) - 1
'        If Fnd_Md(mMd, mPrj, mAnm(J)) Then ss.A 3: GoTo E
'        If Gen_Doc_For1Mod(mF, mMd, pLikNmPrc) Then ss.A 4: GoTo E
'    Next
'    Close #mF
'End If
'
'''GenDoc for Queries into text file {<MdbNam>_Qry.csv}
'If pLikNmq <> "" Then
'    '''Write Header.
'    If Opn_Fil_ForOutput(mF, mFxDocPfx & "_Qry.csv", True) Then ss.A 1: GoTo E
'    Write #mF, "#", "QrySet", "Major#", "MajorName", "Minor#", "Type", "MinorName", "UpdRmk", "Remark", "SQL"
'
'    '''Put all QrySet of given prefix in a Collection (note it is already sorted, so no need to sort)
'    Dim iQry As QueryDef, mNmQsLas$, mNmQsCur$, mAnQs$()
'    If Fnd_AnQs(mAnQs, pLikNmq) Then ss.A 2: GoTo E
'
'    '''Loop the Collection and call <zzGenDoc_For1QrySet>
'    For J = 0 To Sz(mAnQs) - 1
'        If Gen_Doc_For1QrySet(mF, mAnQs(J)) Then ss.A 3: GoTo E
'    Next
'    Close #mF
'End If
'
'Dim mWbDoc As Workbook: If Crt_Wb(mWbDoc, mFxDocPfx & "_Doc.xls", True) Then ss.A 1: GoTo E
'''To reference to Xls
'Call mWbDoc.Application.Vbe.ActiveVBProject.References.AddFromGuid("{4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}", 9, 0)
'
''Merge at most 3 csv into one workbook {<MdbNam>_Doc.xls}
'Dim mWs As Worksheet
'If pLikNmm <> "" Then
'    If Add_WsFmCsv(mWs, mWbDoc, mFxDocPfx & "_Mod.csv", "Modules") Then ss.A 4: GoTo E
'    If Gen_Doc_FmtMod(mWs) Then ss.A 5: GoTo E
'End If
'
'If pLikFrm <> "" Then
'    If Add_WsFmCsv(mWs, mWbDoc, mFxDocPfx & "_Frm.csv", "Forms") Then ss.A 6: GoTo E
'    If Gen_Doc_FmtMod(mWs) Then ss.A 7: GoTo E
'End If
'
'If pLikNmq <> "" Then
'    If Add_WsFmCsv(mWs, mWbDoc, mFxDocPfx & "_Qry.csv", "Queries") Then ss.A 8: GoTo E
'    If Gen_Doc_FmtMod(mWs) Then ss.A 9: GoTo E
'End If
'mWbDoc.Application.Visible = True
'mWbDoc.SaveAs mFxDocPfx & "_Doc.xls", XlFileFormat.xlWorkbookNormal
'Exit Function
'R: ss.R
'E: Gen_Doc = True: ss.B cSub, cMod, "pLikNmm,pLikNmq,pLikFrm", pLikNmm, pLikNmq, pLikFrm
'End Function

'Function Gen_Doc__Tst()
'Close
'If Gen_Doc Then Stop
'End Function

'Function Gen_Doc_QryDpd(pPfxNmq$, Optional pDb As database) As Boolean
''Aim: Gen a Xls "QryDpd" in SdirDoc of 1 worksheet
'Const cSub$ = "Gen_Doc_QryDpd"
'On Error GoTo R
'Dim mDb As database: Set mDb = DbNz(pDb)
'Dim mFfnnQryDpd$: mFfnnQryDpd = Sdir_Doc & "QryDpd"
'Dim mAnq$(): If Fnd_Anq_ByPfx(mAnq, pPfxNmq, mDb) Then ss.A 1: GoTo E
'Dim mFno As Byte
'If Opn_Fil_ForOutput(mFno, mFfnnQryDpd & "_QryDpd.csv") Then ss.A 2: GoTo E
'Dim iQry As DAO.QueryDef
'Write #mFno, "Nmq", "Typ", "DependOn"
'Dim J%, N%, I%
'Dim mAnt$()
'N% = Sz(mAnq)
'For J = 0 To N - 1
'    Set iQry = mDb.QueryDefs(mAnq(J))
'    Dim mSql$: mSql = iQry.Sql
'    If SqsToAnt(mAnt, mSql) Then ss.A 3: GoTo E
'    For I = 0 To Sz(mAnt) - 1
'        Write #mFno, mAnq(J), ToStr_TypQry(iQry.Type), mAnt(I)
'    Next
'Next
'Close #mFno
''Create Doc Xls
'Dim mWb As Workbook: If Crt_Wb(mWb, mFfnnQryDpd & ".xls") Then ss.A 1: GoTo E
'Dim mWs As Worksheet
'If Add_WsFmCsv(mWs, mWb, mFfnnQryDpd & "_QryDpd.csv", "QryDpd") Then ss.A 1: GoTo E
'Call Set_Zoom(mWs, 80)
'mWs.Columns.AutoFit
'Set_Silent
'Dlt_Ws_InWb mWb, "ToBeDelete"
'mWb.Application.Visible = True
'GoTo X
'R: ss.R
'E: Gen_Doc_QryDpd = True: ss.B cSub, cMod, "pPfxNmq,pDb", pPfxNmq, ToStr_Db(pDb)
'X:
'    Set_Silent_Rst
'End Function

'Function Gen_Doc_QryDpd__Tst()
'If Gen_Doc_QryDpd("qryAddTbl") Then Stop
'End Function

'Function Gen_Doc_DbStruct(Optional pInclTbl As Boolean = True, Optional pInclQry As Boolean = True, Optional pInclTypFld As Boolean = False, Optional pCls As Boolean = False) As Boolean
''Aim: Gen a Xls "DbStruct" in SdirDoc of 1 or 2 worksheets.  1 row = 1 object(Tbl or Qry).  1 row = name + list fields.  1 field = Nm(Tnn).
''   T=Byte,Int,Lng,Sng,Dbl,N(Dec),Text,Moment(Date/Time),Y(YesNo)
''   nn is for N,T
'Const cSub$ = "Gen_Doc_DbStruct"
'On Error GoTo R
'Dim mFfnnDbStruct$: mFfnnDbStruct = Sdir_Doc & "DbStruct"
'Dim iFld As DAO.Field
'Dim mFno As Byte
'        Dim mA$
'If pInclTbl Then
'    If Opn_Fil_ForOutput(mFno, mFfnnDbStruct & "_Tbl.csv") Then ss.A 1: GoTo E
'    Dim iTbl As DAO.TableDef
'    For Each iTbl In CurrentDb.TableDefs
'        If IsTbl(iTbl.Name) Then Write #mFno, iTbl.Name, ToStr_TblAtr(iTbl.Attributes), ToStr_Flds(iTbl.Fields, pInclTypFld), Fnd_Prp(iTbl.Name, acTable, "Description")
'    Next
'    Close #mFno
'End If
'If pInclQry Then
'    If Opn_Fil_ForOutput(mFno, mFfnnDbStruct & "_Qry.csv") Then ss.A 2: GoTo E
'    Dim iQry As DAO.QueryDef
'
'    Write #mFno, "Nmqs", "Nmq", "Typ", "Fields", "Desc"
'    For Each iQry In CurrentDb.QueryDefs
'        If IsQry(iQry.Name) Then
'            Select Case iQry.Type
'            Case DAO.QueryDefTypeEnum.dbQSelect _
'                , DAO.QueryDefTypeEnum.dbQSelect _
'                , DAO.QueryDefTypeEnum.dbQCrosstab _
'                , DAO.QueryDefTypeEnum.dbQSQLPassThrough
'                Dim mNmQs$, p%: p = InStr(iQry.Name, "_"): mNmQs = IIf(p = 0, "", Left(iQry.Name, p - 1))
'                Write #mFno, mNmQs, iQry.Name, ToStr_TypQry(iQry.Type), ToStr_Flds(iQry.Fields, pInclTypFld), Fnd_Prp(iQry.Name, acQuery, "Description")
'            End Select
'        End If
'    Next
'    Close #mFno
'End If
''Create Doc Xls
'Dim mWb As Workbook: If Crt_Wb(mWb, mFfnnDbStruct & ".xls") Then ss.A 1: GoTo E
'Dim mWs As Worksheet:
'If pInclTbl Then
'    If Add_WsFmCsv(mWs, mWb, mFfnnDbStruct & "_Tbl.csv", "Tables") Then ss.A 1: GoTo E
'    Call Set_Zoom(mWs, 80)
'    mWs.Columns.AutoFit
'End If
'If pInclQry Then
'    If Add_WsFmCsv(mWs, mWb, mFfnnDbStruct & "_Qry.csv", "Queries") Then ss.A 1: GoTo E
'    Call Set_Zoom(mWs, 80)
'    mWs.Columns.AutoFit
'End If
'Set_Silent
'Dlt_Ws_InWb mWb, "ToBeDelete"
'If Sav_Wb(mWb) Then ss.A 1: GoTo E
'If pCls Then Cls_Wb mWb, False: GoTo X
'mWb.Application.Visible = True
'GoTo X
'R: ss.R
'E: Gen_Doc_DbStruct: ss.B cSub, cMod, "pInclTbl,pInclQry,pInclTypFld,pCls", pInclTbl, pInclQry, pInclTypFld, pCls
'X: Set_Silent_Rst
'End Function
'Function Gen_Doc_FmtMod(pWs As Worksheet) As Boolean
'Const cSub$ = "Gen_Doc_FmtMod"
''mXls.ActiveSheet.Cells(2, 1).Select
''pWs.Application.Calculation = xlCalculationAutomatic
''pWs.Application.ScreenUpdating = True
'On Error GoTo R
'Dim mL&
'Dim mLasRow&
'Dim mRange As Range
'With pWs
'    Debug.Print "Format Ws[" & pWs.Name & "] --- Outline ...."
'    .Outline.SummaryRow = xlAbove
'    'Loop all rows from row 3 to set outlinelevel
'    mLasRow = .Cells.SpecialCells(xlCellTypeLastCell).Row
'    For mL = 3 To mLasRow
'        If mL Mod 500 = 0 Then Debug.Print Fmt_Str("Set outline & hyperline for line: {0}  {1}", mL, mLasRow)
'        Dim mOL As Byte: mOL = .Cells(mL, 1).Value
'        .Rows(mL).OutlineLevel = mOL
'
'        If mOL >= 3 Then
'            '''Set Hyperlinks in column D (The Line#)
'            Set mRange = .Cells(mL, 4)
'            If Not IsNull(mRange.Value) Then
'                mRange.Hyperlinks.Add mRange, "", .Name & "!C" & mL
'
'                '''Set RED if it is #Chk# or #Skip#
'                Set mRange = .Cells(mL, 4)
'                If InStr(mRange.Value, "#Check") > 0 Or InStr(mRange.Value, "#Skip#") Then mRange.Font.Color = RGB(255, 0, 0)
'            End If
'        End If
'    Next
'    'Delete column A & Format the columns
'    .Columns("A").Delete
'
'    .Columns("D:G").Font.Name = "Courier New"
'    .Columns("A:G").ColumnWidth = 4
'    .Columns("C").ColumnWidth = 5
'    .Outline.ShowLevels 3
'    .Outline.ShowLevels 2
'End With
''Add module:
'Dim mVbCmp As VBIDE.VBComponent: If Fnd_VbCmp_FmWs(mVbCmp, pWs) Then ss.A 1: GoTo E
'''Find mVBCmp as Sheet1 & mXlsCode by calling <Fnd.StringFmMod(<Nmm>,<NmPrc>)
'
'Dim mXlsCode$: If Fnd_ResStr(mXlsCode, "GenDoc_FmtMod", True) Then ss.A 2: GoTo E
'''Add code to the worksheet
'
'mVbCmp.CodeModule.AddFromString mXlsCode
'pWs.Application.ActiveWindow.Zoom = 80
'Exit Function
'R: ss.R
'E: Gen_Doc_FmtMod = True: ss.B cSub, cMod, "pWs", ToStr_Ws(pWs)
'End Function
'Private Function Gen_Doc_FmtQry(pWs As Worksheet) As Boolean
'Const cSub$ = "Gen_Doc_FmtQry"
'On Error GoTo R
'Dim mL&
'Dim mLasRow&
'Dim mRange As Range
''mXls.ActiveSheet.Cells(2, 1).Select
''pWs.Application.Calculation = xlCalculationAutomatic
''pWs.Application.ScreenUpdating = True
'With pWs
'    mLasRow = .Cells.SpecialCells(xlCellTypeLastCell).Row
'    ''Loop all rows from row 2 to set outlinelevel
'    .Outline.SummaryRow = xlAbove
'    For mL = 2 To mLasRow
'        .Rows(mL).OutlineLevel = .Range("A" & mL).Value
'    Next
'    ''Delete first column and set column width
'    .Columns("$A").Delete
'
'    ''Loop all rows from row 2 again to set hyperlinks
'    For mL = 2 To mLasRow
'        ''  If Lvl2 or Lvl3, Set hyperlink to [<UpdRmk>]
'        Select Case .Rows(mL).OutlineLevel
'        Case 2, 3
'            Set mRange = .Range("G" & mL)
'            mRange.Hyperlinks.Add mRange, "", .Name & "!" & mRange.Address
'        End Select
'        ''  If Lvl3, Set hyperlink to [<Min>]
'        If .Rows(mL).OutlineLevel = 3 Then
'            Set mRange = .Range("D" & mL)
'            mRange.Hyperlinks.Add mRange, "", .Name & "!" & mRange.Address
'
'            ''  If Typ is select, union or crosstable, set hyperlinek to [<Typ>]
'            Select Case .Range("E" & mL).Value
'            Case "Select", "SetOperation", "Crosstab"   '"SetOperation means union
'                Set mRange = .Range("E" & mL)
'                mRange.Hyperlinks.Add mRange, "", .Name & "!E" & mL
'            End Select
'        End If
'    Next
'    ''Somemore formatting
'    '''Format column A:F = 4
'    '''Format column G   = 8                <UpdRmk>
'    '''Format column E   = 10
'    '''Format column F   = 40
'    '''Format column G   = 10
'    '''Format column I   = 100 & WrapText
'    .Columns("A:F").ColumnWidth = 4
'    .Columns("G").ColumnWidth = 8
'    .Columns("E").ColumnWidth = 10
'    .Columns("F").ColumnWidth = 40
'    .Columns("I").ColumnWidth = 80
'    .Columns("I").WrapText = True
'    '''HorizontalAlign column B,D = Center
'    .Columns("B").HorizontalAlignment = xlCenter
'    .Columns("D").HorizontalAlignment = xlCenter
'    ''' ...
'    With .Rows("1:" & mLasRow)
'        .AutoFit
'        .VerticalAlignment = xlTop
'    End With
'    .Outline.ShowLevels 3
'    .Outline.ShowLevels 2
'End With
'
'''Find <mVBCmp> & <mXlsCode> by calling <Fnd.StringFmMod(<Nmm>,<NmPrc>)
'Dim mVbCmp As VBComponent: If Fnd_VbCmp_FmWs(mVbCmp, pWs) Then ss.xx 1, cSub, cMod: Exit Function
'Dim mXlsCode$: If Fnd_ResStr(mXlsCode, "zzGenDoc_FmtQry", True) Then ss.xx 2, cSub, cMod: Exit Function
'
'''Add code to the worksheet
'mVbCmp.CodeModule.AddFromString mXlsCode
'pWs.Application.ActiveWindow.Zoom = 80
'Exit Function
'R: ss.R
'E: Gen_Doc_FmtQry = True: ss.A cSub, cMod, ToStr_Ws(pWs)
'End Function
'Private Function Gen_Doc_For1Mod(pFno As Byte, pMd As CodeModule, pLikNmPrc$) As Boolean
'Const cSub$ = "Gen_Doc_For1Mod"
''Aim: GenDoc of all proc of prefix {pPfxProc} of given module {pNmm} of type {pAcObjNam: Form/Module} to {pFno} in following format
'''1,<MdbNam_Full>
'''2,,<Proc>
'''3,,,<Line>,'Remark
'''4,,,<Line>,Code
'''4,,,<Line>,''Remark
'''5,,,<Line>,Code
'''5,,,<Line>,'''Remark
'''6,,,<Line>,Code
'''6,,,<Line>,''''Remark
'''7,,,<Line>,Code
''Notes: In order to access the procedures of the given {pNmm}, it will be openned and then closed
''==Start
'On Error GoTo R
''Open the given module {pNmm} in <pMd>
'If TypeName(pMd) = "Nothing" Then Exit Function
'
''Write Lvl1
'Dim mNmm$: mNmm = ToStr_Md(pMd)
'Write #pFno, 1, mNmm
'Debug.Print mNmm
''Loop each procedure {iNmPrc} of Prefix {pLikNmPrc} of <pMd> to write lines to {pFno}
'Dim iPrc, mCurLvl As Byte, mTyp&, iLinNo
'Dim mAnPrc$(): If Fnd_AnPrc_ByMd(mAnPrc, pMd, pLikNmPrc) Then ss.A 1: GoTo E
'Dim J%
'For J = 0 To Sz(mAnPrc) - 1
'    Dim iNmPrc$, iPrcBeg$, iPrcEnd$
'
'    If Brk_Str_To3Seg(iNmPrc, iPrcBeg, iPrcEnd, mAnPrc(J), ":") Then ss.xx 1, cSub, cMod: Exit Function
'    Debug.Print Chr(9) & iNmPrc
'
'    '''Write Lvl2 of {iNmPrc} and first line of Lvl3
'    Write #pFno, 2, , iNmPrc
'    Write #pFno, 3, , , iPrcBeg, pMd.Lines(iPrcBeg, 1)
'    mCurLvl = 3
'
'    ''Loop all lines {iLine} {iLineTrim} within the procedure {iNmPrc} according to <mProcLen>, <mProcLen>
'    For iLinNo = iPrcBeg + 1 To iPrcEnd
'        Dim iLine$:        iLine = pMd.Lines(iLinNo, 1)
'        Dim iLineTrim$:    iLineTrim = Trim(iLine)
'
'        '''Set mCurLvl
'        Dim mIsRmk As Boolean: mIsRmk = True
'
'        If Left(iLineTrim, 5) = "'''''" Then
'                                                mCurLvl = 7
'        ElseIf Left(iLineTrim, 4) = "''''" Then
'                                                mCurLvl = 6
'        ElseIf Left(iLineTrim, 3) = "'''" Then
'                                                mCurLvl = 5
'        ElseIf Left(iLineTrim, 2) = "''" Then
'                                                mCurLvl = 4
'        ElseIf Left(iLineTrim, 1) = CtSngQ Then
'                                                mCurLvl = 3
'        Else
'                                                mIsRmk = False
'        End If
'
'        Write #pFno, mCurLvl + IIf(mIsRmk, 0, 1), , , iLinNo, iLine
'    Next
'Next
'Exit Function
'R: ss.R
'E: Gen_Doc_For1Mod = True: ss.B cSub, cMod, "pFno,pMd,pLikNmPrc", pFno, ToStr_Md(pMd), pLikNmPrc$
'End Function
'Private Function Gen_Doc_For1QrySet(pFno As Byte, QryNms$) As Boolean
'Const cSub$ = "Gen_Doc_For1QrySet"
'On Error GoTo R
'Dim iQry As DAO.QueryDef, mTyp$, mLasMaj$, mL As Byte
'mLasMaj = "??"
''Aim: Write documentation for all queries in <QryNmLst> of prefix as in <QryNms> to file <#pF>
''==Start
''Write Lvl 1
'Write #pFno, 1, QryNms
''Loop each iNmq in <QryNmLst>
'Dim mAnq$(): If Fnd_Anq_ByNmQs(mAnq, QryNms) Then ss.A 1: GoTo E
'Dim mDQry As New d_Qry
'With mDQry
'    Dim J%: For J = 0 To Sz(mAnq) - 1
'        Dim iNmq$: iNmq = mAnq(J)
'        Set iQry = CurrentDb.QueryDefs(iNmq)
'        If mDQry.Brk_Nmqs(iQry.Name) Then ss.A 1: GoTo E
'        If QryNms <> .NmQs Then ss.A 1: GoTo E
'        mTyp = ToStr_TypQry(iQry.Type)
'        Dim mRmk$:    mRmk = ""
'        On Error Resume Next
'        mRmk = iQry.Properties("Description").Value
'        If mLasMaj <> .Maj Then
'            If .Min <> 0 Then ss.A 1, "There is no Min Step 0", , "The Query Set Nam,mLasMaj,mMaj", QryNms, mLasMaj, .Maj: GoTo E
'            If iQry.Type <> DAO.QueryDefTypeEnum.dbQSelect Then ss.A 2, "The query of minor step 0 must be select query", , "The Query,Query Type(DAO.QueryDefTypeEnum)", iQry.Name, iQry.Type: GoTo E
'            'Write Lvl 2 if needed
'            Write #pFno, 2, , .Maj, mID$(iNmq, mL + 7), , , , "UpdRmk", mRmk ' Mid$(iNmq,mL+7) is "primary file in current step"
'            mRmk = ""
'            mLasMaj = .Maj
'        End If
'        On Error GoTo 0
'        'Write Lvl 3
'        Write #pFno, 3, , , , .Min, .Typ, mID$(iNmq, mL + 7), "UpdRmk", .Des
'        'Write Lvl 3
'        Write #pFno, 4, , , , , , , , , Replace(iQry.Sql, Chr(13), "")
'Nxt:
'    Next
'End With
'Exit Function
'R: ss.R
'E: Gen_Doc_For1QrySet = True: ss.B cSub, cMod, "pFno,QryNms", pFno, QryNms
'End Function
'Function Gen_Doc_Template() As Boolean
'Const cSub$ = "Gen_Doc_Template"
'Dim mFno As Byte:       mFno = FreeFile
'Dim mFfnCsv$:  mFfnCsv = Sdir_Doc & Sffn_Cur & "_Doc(ForTemplate).csv"
'Dim mDirTp$:   mDirTp = Sdir_Tp
'Dim mAyFn$(): If Fnd_AyFn(mAyFn, mDirTp) Then ss.A 1: GoTo E
'If Sz(mAyFn) = 0 Then ss.A 1, "There is not Xls file in Template Dir", , "mDirTp", mDirTp: GoTo E
'Open mFfnCsv For Output As #mFno
'Dim iFn
'For Each iFn In mAyFn
'    Dim mWb As Workbook: Set mWb = gXls.Workbooks.Open(mDirTp & iFn, False, , , , , True)
'    With mWb
'        Write #mFno, mWb.Name, , , , mWb.FullName
'        If mWb.PivotCaches.Count > 0 Then
'            Write #mFno, , "PivotCaches.Count(" & mWb.PivotCaches.Count & ")"
'            Dim iPc As Excel.PivotCache
'            For Each iPc In .PivotCaches
'                Write #mFno, , , ;: Print #mFno, ToStr_Pc(iPc)
'            Next
'        End If
'        Dim iWs As Worksheet
'        For Each iWs In mWb.Worksheets
'            If iWs.PivotTables.Count > 0 Then
'                Write #mFno, , "PivotTables.Count(" & iWs.PivotTables.Count & ") Ws(" & iWs.Name & ")"
'                Dim iPt As PivotTable
'                For Each iPt In iWs.PivotTables
'                    Write #mFno, , , ;: Print #mFno, ToStr_Pt(iPt)
'                Next
'            End If
'        Next
'        For Each iWs In mWb.Worksheets
'            If iWs.QueryTables.Count > 0 Then
'                Write #mFno, , "QueryTables.Count(" & iWs.QueryTables.Count & ") Ws(" & iWs.Name & ")"
'                Dim iQt As Excel.QueryTable
'                For Each iQt In iWs.QueryTables
'                    Write #mFno, , , ;: Print #mFno, ToStr_Qt(iQt)
'                Next
'            End If
'        Next
'        .Close False
'    End With
'Next
'Close #mFno
''Format the csv to xls
'Dim mXls As New Excel.Application
'Set mWb = mXls.Workbooks.Open(mFfnCsv)
'Dim mWs As Worksheet: Set mWs = mWb.Worksheets(1)
'If WsFmtOL(mWs, 3) Then ss.A 2: GoTo E
'mWs.Columns(3).ColumnWidth = 40
'mWs.Columns(4).ColumnWidth = 15
'Dlt_Fil Left(mFfnCsv, Len(mFfnCsv) - 4) & ".xls"
'mWb.SaveAs Left(mFfnCsv, Len(mFfnCsv) - 4) & ".xls"
'mXls.Visible = True
'Exit Function
'R: ss.R
'E: Gen_Doc_Template = True: ss.B cSub, cMod, ""
'End Function
'Function Gen_Prn_ByFDF() As Boolean
'Const cSub$ = "Gen_Prn_ByFDF"
''Aim: Allow user to pick a *.FDF in a directory to create a *.prn by *.xls.
'On Error GoTo R
'Dim mFfnFdf$:  If Fnd_Ffn(mFfnFdf, "c:\", "*.FDF") Then GoTo E
'Dim mF As Byte: If Opn_Fil_ForInput(mF, mFfnFdf) Then ss.A 1: GoTo E
'
''Find AyCnoWidth(1 to n) from given FDF
'Dim AyCnoWidth() As Byte ' ColWidth of as column as described in the given FDF
'Dim aDec() As Byte      ' # of Decimal Place for numerice column as described in the given FDF
'Dim mL$, iL As Byte, aa$(), aB$()
''Skip 2 lines
'Line Input #mF, mL
'Line Input #mF, mL
'Line Input #mF, mL
'While Not EOF(mF)
'    Line Input #mF, mL
'    iL = iL + 1
'    ReDim Preserve AyCnoWidth(1 To iL), aDec(1 To iL)
'    aa = Split(mL, " ")
'    If aa(0) <> "PCFL" Then Stop
'    Select Case aa(2)
'    Case "1": AyCnoWidth(iL) = aa(3)
'    Case "2":
'                aB = Split(aa(3), "/")
'                If LBound(aB) <> 0 Then Stop
'                Select Case UBound(aB)
'                Case 0
'                Case 1: aDec(iL) = aB(1)
'                Case Else: Stop
'                End Select
'                AyCnoWidth(iL) = aB(0)
'    Case Else
'        ss.A 2, "After [PCFL], it must be 1 or 2, but it is now [" & aa(0) & "]": GoTo E
'    End Select
'Wend
'Close #mF
'If False Then
'    Dim J As Byte
'    For J = 1 To UBound(AyCnoWidth)
'        Debug.Print ToStr_LpAp(CtComma, "Column,Width", J, AyCnoWidth(J))
'    Next
'    Stop
'End If
'
''Open Xls and gen prn
'Dim mWs As Worksheet
'Dim mFfnn$: mFfnn = Left(mFfnFdf, Len(mFfnFdf) - 4)
'Dim mWb As Workbook: If Opn_Wb(mWb, mFfnn & ".xls", True) Then ss.A 3: GoTo E
'Set mWs = mWb.Worksheets(1)
'mWs.Rows(1).Delete
'For J = 1 To UBound(AyCnoWidth)
'    mWs.Columns(J).ColumnWidth = AyCnoWidth(J)
'    If aDec(J) > 0 Then
'        mWs.Columns(J).NumberFormat = "0." & String(aDec(J), "0")
'    End If
'Next
'mWb.Application.DisplayAlerts = False
'mWb.SaveAs mFfnn & ".prn", XlFileFormat.xlTextPrinter
'mWb.Close False
'mWb.Application.DisplayAlerts = True
'R: ss.R
'E: Gen_Prn_ByFDF = True: ss.B cSub, cMod
'End Function
'Function Gen_Rpt(pNmRptSht$, pNmSess$, Optional pLm$) As Boolean
'Const cSub$ = "Gen_Rpt"
'Dim mSkp_Download As Boolean
'Dim mSkp_Download_Skip:     mSkp_Download_Skip = False
'Dim mSkp_GenXls_AllData:    mSkp_GenXls_AllData = False
'Dim mSkp_GenXls_EachData:   mSkp_GenXls_AllData = False
'Dim mSkp_RunQry:            mSkp_GenXls_AllData = False:
'Dim mChk_AllDtaXls As Boolean:  mChk_AllDtaXls = False
'Dim mChk_Download As Boolean:   mChk_Download = False
'Dim mChk_EachDtaXls As Boolean: mChk_EachDtaXls = False
'Dim mChk_Prm As Boolean:        mChk_Prm = False
'
''Aim: Generate one or more Xls files in .\Output\{NmSess}\{rptOFmtStr_FnTo}.xls in 3 steps Download RunQry GenXls
''Assumes:
'''Tables & Queries for Download Parameters
'''tblPrm_{Nmrptsht}  must exist and contains: NmSess, DownloadNam, Env
'''mstEnv, mstLib, mstIP must exist
'''tblRpt                must exist and contains a record of Nmrptsht='{pNmRptSht}'
'''qRpt{Nmrptsht} must exists, which join tblPrm_{Nmrptsht}, mstEnv, mstLib, mstIP, tblRpt to form a DownloadPrmRs,
'''                      which will always have 3 columns
'''
'''Template:
'''Location = {Home}\WorkingDir\Templates\{Nmrptsht}_Template.xls
'''Notes    : It can contain macro string of {NmRpt} {Nmrptsht} {NmSess} {NmDta} in Chart Title or Page Headers
'''
'''DTF (Optional)
'''Location (DTF)         = {Home}\WorkingDir\DTF_{Nmrptsht}\src\*.DTF
'''Location (Empty Xls) = {Home}\WorkingDir\DTF_{Nmrptsht}\EmptyXls\*.xls
'''Notes                  : DTF name and the download target Xls file name must be the same, otherwise, it will look up from EmptyXls
'''                         to copy.
''==Start==
'''Set mIsBatchMode if {pXls} is not given
'ss.xx 1, cSub, cMod, eTrc, "Start", "pNmRptSht, pNmsess", pNmRptSht, pNmSess
'
''Preparation
'''Verify pNmRpt, pNmSess
''''Record of {NmRpt} should exist in tblRpt
'Dim mCnt&: If Fnd_RecCnt_ByNmtq(mCnt, "tblRpt", "Nmrptsht='" & pNmRptSht & CtSngQ) Then ss.A 2, "No records in tblRpt": GoTo E
'
'''Get report parameter from tblRpt in variables mP.* & mP.Each*
'
'Dim mP As tRpt
'If Fnd_TypPrmRpt(mP, pNmRptSht) Then ss.A 1, "Given report not define in tblRpt": GoTo E
'
'''#Chk# Prm
'If mChk_Prm Then
'    Shw_Dbg cSub, cMod, "Check calling param", "pNmRptSht, pNmSess, mP.NmRpt , mP.FmtStr_FnTo, mP.LnwsRmv , mP.HidePfLst_ThisNmSess, mP.HidePfLst_ThisSess, mP.HidePfLst_OtherSess, mP.NmDta, mP.EachNmFld, mP.EachHidePfLst_ThisSess, mP.EachHidePfLst_OtherSess", _
'        pNmRptSht, pNmSess, mP.NmRpt, mP.FmtStr_FnTo, mP.LnwsRmv, mP.HidePfLst_ThisNmSess, mP.HidePfLst_ThisSess, mP.HidePfLst_OtherSess, mP.NmDta, mP.EachNmFld, mP.EachHidePfLst_ThisSess, mP.EachHidePfLst_OtherSess
'    Stop
'End If
'Dim mMsg$:
'mMsg = Fmt_Str("Generate [{0}] for [{1}]?", mP.NmRpt, pNmSess)
'If Not Fct.Start(mMsg, "Generate Report?") Then Exit Function
'
'Dim mLm$: mLm = Add_Str(pLm, "Date=" & Format(Date, "YYYY_MM_DD"))
'mLm = mLm & ",NmDta=" & mP.NmDta
'mLm = mLm & ",NmSess=" & pNmSess
'mLm = mLm & ",NmRptSht=" & pNmRptSht
'mLm = mLm & ",NmRpt=" & mP.NmRpt
'mLm = mLm & ",MGIWeekNum=Wk" & Fct.MGIWeekNum(Date)
'
''RunQry qry{NmRptSht} #Skip#
'If mSkp_RunQry Then
'    Stop
'Else
'    ''RunQry
'    If mSkp_Download Then
'        If Run_Qry("qry" & pNmRptSht, , , , mLm) Then ss.A 1: GoTo E
'    Else
'        If Run_Qry("qry" & pNmRptSht, , , , mLm, , True) Then ss.A 1: GoTo E
'    End If
'End If
'
''Prepare GenXls for AllData and/or EachData
'''Set mFmtStr_FnTo=p.rptOFmtStr_FnTo if given, else set to {NmSess} {NmRpt} {MGIWeekNum}@{Date}{NmDta}
'Dim mFmtStr_FnTo$
'If mP.FmtStr_FnTo = "" Then
'    mFmtStr_FnTo = "{NmSess} {NmRpt} {MGIWeekNum}@{Date}{NmDta}.xls"
'Else
'    mFmtStr_FnTo = mP.FmtStr_FnTo
'End If
'
'''Prepare <mCollHdrMacro>: It always has 2 variables: NmSess & instNam
'Dim mCollHdrMacro As New VBA.Collection
'mCollHdrMacro.Add "NmSess=" & pNmSess
'mCollHdrMacro.Add "", "NmDta"
'
'Dim mDocPrp As tDocPrp
''With mDocPrp
''    .namRpt = mP.NmRpt
''    .namRptShort = pNmRptSht
''    .namSess = pNmSess
''    .ExtraPrm = ToStr_Coll(Fnd.ExtraPrm)
''End With
'
''GenXls_EachData #Skip#
'Dim mFfnTo$
'Dim mFfnFm$: mFfnFm = Sffn_Tp(pNmRptSht)
'If mSkp_GenXls_EachData Then
'    Stop
'Else
'    ''Gen Xls For each {Data} if <p.eachSql> is given
'    If mP.EachSql <> "" Then
'        ''Loop <p.eachSql>
'        Dim mRsEach As DAO.Recordset
'        Set mRsEach = CurrentDb.OpenRecordset(mP.EachSql)
'        Dim mAm() As tMap: mAm = Get_Am_ByLm(mLm)
'        With mRsEach
'            While Not .EOF
'                ''Set <mCollHdrMacro>
'                mCollHdrMacro.Remove "NmDta"
'                mCollHdrMacro.Add "NmDta=" & .Fields(mP.EachNmFld).Value, "NmDta"
'                mDocPrp.NmData = .Fields(mP.EachNmFld).Value
'
'                ''Set <mFfnTo> from mLp, mAm
'                If Set_Am_ByF1F2(mAm, "NmDta", .Fields(mP.EachNmFld).Value) Then ss.A 1: GoTo E
'                mFfnTo = Sdir_RptSess(pNmSess) & Fmt_Str_ByAm(mFmtStr_FnTo, mAm)
'
'                ''Gen Xls each data
'                If Gen_Xls(mFfnFm, mFfnTo, _
'                        pKeepWbOpnAndNotSav:=True, _
'                        pCollMacro:=mCollHdrMacro, _
'                        pLnWsRmv:=mP.LnwsRmv, _
'                        pLExpr:=Fmt_Str("[{0}]='{1}'", mP.EachNmFld, .Fields(mP.EachNmFld))) Then _
'                      ss.A 1: GoTo E
'
'                ''Hide Pf if needed
'                If mP.HidePfLst_ThisNmSess = pNmSess Then
'                    If mP.EachHidePfLst_ThisSess <> "" Then Call Set_Pf_OfWb(gXls.Workbooks(1), mP.EachHidePfLst_ThisSess)
'                Else
'                    If mP.EachHidePfLst_OtherSess <> "" Then Call Set_Pf_OfWb(gXls.Workbooks(1), mP.EachHidePfLst_OtherSess)
'                End If
'
'                ''#Chk#: EachDataXls file generated
'                If mChk_EachDtaXls Then
'                    gXls.Visible = True
'                    Stop
'                    gXls.Visible = False
'                End If
'
'                Set_DocPrp gXls.Workbooks(1), mDocPrp
'                ''Save & Close
'                Call gXls.Workbooks(1).Close(True)
'                .MoveNext
'            Wend
'            .Close
'        End With
'    End If
'End If
'
''GenXls_AllData #Skip#
'If mSkp_GenXls_AllData Then
'    Stop
'Else
'    ''Set <mCollHdrMacro>
'    mCollHdrMacro.Remove "NmDta"
'    mCollHdrMacro.Add "NmDta=" & mP.NmDta, "NmDta"
'
'    ''Set <mFfnTo> from <mLp> & <mAm>
'    If Set_Am_ByF1F2(mAm, "NmDta", mP.NmDta) Then ss.A 1: GoTo E
'    mFfnTo = Sdir_RptSess(pNmSess) & Fmt_Str_ByAm(mFmtStr_FnTo, mAm)
'
'    ''Gen Xls for all data
'    If Gen_Xls(mFfnFm, mFfnTo, _
'        pKeepWbOpnAndNotSav:=True, _
'        pLnWsRmv:=mP.LnwsRmv, _
'        pCollMacro:=mCollHdrMacro) Then _
'          ss.A 1, "Error in gen all data report": GoTo E
'
'    ''Hide Pf if needed
'    If mP.HidePfLst_ThisNmSess = pNmSess Then
'        If mP.HidePfLst_ThisSess <> "" Then Call Set_Pf_OfWb(gXls.Workbooks(1), mP.HidePfLst_ThisSess)
'    Else
'        If mP.HidePfLst_OtherSess <> "" Then Call Set_Pf_OfWb(gXls.Workbooks(1), mP.HidePfLst_OtherSess)
'    End If
'
'    ''#Chk#: AllDataXls file generated
'    If mChk_AllDtaXls Then
'        gXls.Visible = True
'        Stop
'    End If
'
'    ''Set Document Properties
''    With mDocPrp
''        .NmRpt = mP.NmRpt
''        .NmRptSht = pNmRptSht
''        .NmSess = pNmSess
''        .NmData = mP.NmDta
''        .ExtraPrm = ToStr_Coll(Fnd.ExtraPrm)
''    End With
'    Set_DocPrp gXls.Workbooks(1), mDocPrp
'    ''Save & Close
'    gXls.Workbooks(1).Close True
'End If
'ss.xx 1, cSub, cMod, eTrc, "End", "mP.NmRpt , pNmSess", mP.NmRpt, pNmSess
'Fct.Done
'Exit Function
'R: ss.R
'E: Gen_Rpt = True: ss.B cSub, cMod, "pNmRptSht,pNmsess", pNmRptSht, pNmSess
'X: Clr_Sts
'End Function
'Function Gen_Rpt_ByBatch() As Boolean
'Const cSub$ = "Gen_Rpt_ByBatch"
''Aim: Use CtCommand() as "{Nmrptsht},{NmSess} to generate report
''==Start
''Verify CtCommand() is a valid NmSess or not
'On Error GoTo R
'Dim mNmrptsht$, mNmSess$: If Fnd_SegFmCmd_2(mNmrptsht, mNmSess) Then ss.A 1: GoTo E
'ss.xx 2, cSub, cMod, eTrc, "Start", "mNmrptsht, mNmSess", mNmrptsht, mNmSess
'IsBch = True
'Dim mA$
'If Gen_Rpt(mNmrptsht, mNmSess) Then
'    mA = "Gen Report By Batch End - Fail"
'Else
'    mA = "Gen Report By Batch End - OK"
'End If
'ss.xx 2, cSub, cMod, eTrc, "Start", "mNmrptsht,mNmSess,Status", mNmrptsht, mNmSess, mA
'Exit Function
'R: ss.R
'E: Gen_Rpt_ByBatch = True: ss.B cSub, cMod, "mNmrptsht,mNmSess", mNmrptsht, mNmSess
'End Function
'Function Gen_TxtFil_ByAy(pFfnFm$, pFfnTo$, pAyK$(), pAyV$()) As Boolean
'Dim mFmFNo As Byte: mFmFNo = FreeFile: Open pFfnFm For Input As #mFmFNo
'Dim mToFNo As Byte: mToFNo = FreeFile: Open pFfnTo For Output As #mToFNo
'While Not EOF(mFmFNo)
'    Dim mLine$: Line Input #mFmFNo, mLine
'    Print #mToFNo, Fmt_Str_ByAyKV(mLine, pAyK, pAyV)
'Wend
'Close #mFmFNo, #mToFNo
'End Function
'Function Gen_TxtFil_ByMacroFil(pFfnFm$, pFfnTo$, pFfnMacro$) As Boolean
'Const cSub$ = "Gen_TxtFil_ByMacroFil"
''Aim: Build a <pFfnTo> from a template <pFfnFm> with referring <pFfnMacro>
'''pFfnFm   : a full path file name of text file contains some substring to be replaced.  The substring is in {<<key>>} format
'''pFfnTo   : a full path file name of text file after macros substitue by using pFfnMacro.  The result.
'''pFfnMacro: a full path file name contain a list of <<key>>=<<value>>
''==Start
''Read <pFfnMacro> into mMacroList
'Dim mAm() As tMap: If Read_MacroFil(mAm, pFfnMacro) Then ss.A 1: GoTo E
'Gen_TxtFil_ByMacroFil = Gen_TxtFil_ByAm(pFfnFm, pFfnTo, mAm)
'Exit Function
'R: ss.R
'E: Gen_TxtFil_ByMacroFil = True: ss.B cSub, cMod, "pFfnFm,pFfnTo,pFfnMacro", pFfnFm, pFfnTo, pFfnMacro
'End Function
'Function Gen_TxtFil_ByAm(pFfnFm$, pFfnTo$, pAm() As tMap) As Boolean
'Const cSub$ = "Gen_TxtFil_ByAm"
'On Error GoTo R
'Dim mFmFNo As Byte: If Opn_Fil_ForInput(mFmFNo, pFfnFm) Then ss.A 1: GoTo E
'Dim mToFNo As Byte: If Opn_Fil_ForOutput(mToFNo, pFfnTo) Then ss.A 2: GoTo E
'While Not EOF(mFmFNo)
'    Dim mLine$: Line Input #mFmFNo, mLine
'    Print #mToFNo, Fmt_Str_ByAm(mLine, pAm)
'Wend
'GoTo X
'R: ss.R
'E: Gen_TxtFil_ByAm = True: ss.B cSub, cMod, "pFfnFm,pFfnTo,pAm", pFfnFm, pFfnTo, ToStr_Am(pAm)
'X:
'    Close #mFmFNo, #mToFNo
'End Function
'Function Gen_TxtFil_ByStrAndRsMacro(pFmtStr$, pFfnTo$, pRsMacro As DAO.Recordset) As Boolean
'Const cSub$ = "Gen_TxtFil_ByStrAndRsMacro"
'Dim FnoTo As Byte: If Opn_Fil_ForOutput(FnoTo, pFfnTo) Then ss.A 1: GoTo E
'Dim Am() As tMap: If RsDic(Am, pRsMacro) Then ss.A 2: GoTo E
'Dim J As Byte
'For J = 0 To Siz_Am(Am) - 1
'    With Am(J)
'        .F2 = Replace(Replace(.F2, Chr(10), " "), Chr(13), " ")
'    End With
'Next
'Print #FnoTo, Fmt_Str_ByAm(pFmtStr, Am)
'GoTo X
'R: ss.R
'E: Gen_TxtFil_ByStrAndRsMacro = True: ss.B cSub, cMod, "pFmtStr,pFfnTo,pRsMacro", pFmtStr, pFfnTo, ToStr_Rs_NmFld(pRsMacro)
'X: Close #FnoTo
'End Function
'Function Gen_TxtFil_ByWs(pFfnTo$, pWs As Worksheet, Optional OvrWrt As Boolean = False) As Boolean
''Aim:   Gen <pFfnTo> from a <pWs>
'Const cSub$ = "Gen_TxtFil_ByWs"
'Dim mFno As Byte: If Opn_Fil_ForOutput(mFno, pFfnTo, OvrWrt) Then ss.A 1: GoTo E
'Dim iRno&
'For iRno = 1 To pWs.Cells.SpecialCells(xlCellTypeLastCell).Row
'    Print #mFno, pWs.Cells(iRno, 1).Value
'Next
'GoTo X
'R: ss.R
'E: Gen_TxtFil_ByWs = True: ss.B cSub, cMod, ""
'X: Close #mFno
'End Function
'Function Gen_Ws(pWb As Workbook, p As tGenWs) As Boolean
'Const cSub$ = "Gen_Ws"
''Create pivot table
'On Error GoTo R
'Dim mNmPt$: mNmPt = IIf(p.Pt_Nam = "", "PivotTable1", p.Pt_Nam)
'Dim mWs As Worksheet: If Add_Ws(mWs, pWb, p.WsNmNew) Then ss.A 1: GoTo E
'StsShw Fmt_Str("Build Wb[{0}] Ws[{1}] Cache[{2}]", pWb.Name, p.WsNmNew, p.Pt_Sqs)
'With pWb.PivotCaches.Add(SourceType:=xlExternal)
'    .Connection = CnnStr_Mdb(CurrentDb.Name)
'    .CtCommandType = xlCmdSql
'    .CtCommandText = p.Pt_Sqs
'    .MaintainConnection = True
'    .CreatePivotTable _
'        TableDestination:=Fmt_Str("'[{0}]{1}'!R1C1", pWb.Name, p.WsNmNew), _
'        TableName:=mNmPt, _
'        DefaultVersion:=xlPivotTableVersion10
'End With
'StsShw Fmt_Str("Build Wb[{0}] Ws[{1}] PivotTable[{2}]", pWb.Name, p.WsNmNew, mNmPt)
'
'Dim mPivotRows$(): mPivotRows = Split(p.PivotRows, CtComma)
'Dim mPivotColumns$(): mPivotColumns = Split(p.PivotColumns, CtComma)
'Dim mPivotData$(): mPivotData = Split(p.PivotData, CtComma)
'Dim J%, mPt As Excel.PivotTable, mPf As Excel.PivotField
'Set mPt = mWs.PivotTables(mNmPt)
'With mPt
'    For J = Sz(mPivotColumns) - 1 To 0 Step -1
'        Set mPf = .PivotFields(mPivotColumns(J))
'        With mPf
'            .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
'            .Orientation = xlColumnField
'            .Position = 1
'            .AutoSort xlAscending, mPivotColumns(J)
'        End With
'    Next
'    For J = Sz(mPivotColumns) - 1 To 0 Step -1
'        Set mPf = .PivotFields(mPivotRows(J))
'        With mPf
'            .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
'            .Orientation = xlRowField
'            .Position = 1
'            .AutoSort xlAscending, mPivotRows(J)
'        End With
'    Next
'    For J = 0 To Sz(mPivotData)
'        Call .AddDataField( _
'            .PivotFields(mPivotData(J)), _
'            " " & mPivotData(J), _
'            xlSum)
'    Next
'    .RowGrand = p.RowGrand
'    .ColumnGrand = p.ColumnGrand
'End With
''Format it
'''ColWidth_Default & ColWidth
'StsShw Fmt_Str("Build Wb[{0}] Ws[{1}] Formatting .....", pWb.Name, p.WsNmNew)
'mWs.Columns.ColumnWidth = p.ColWidth_Default
'Dim mColWidth$(): mColWidth = Split(p.ColWidth, CtComma)
'Dim iCno As Byte
'If UBound(mColWidth) >= LBound(mColWidth) Then
'    For J = LBound(mColWidth) To UBound(mColWidth)
'        iCno = iCno + 1
'        If Val(mColWidth(J)) >= 1 Then mWs.Columns(iCno).ColumnWidth = mColWidth(J)
'    Next
'End If
'''RowHeight
'Dim mRowHeight$(): mRowHeight = Split(p.RowHeight, CtComma)
'Dim iRno&
'If UBound(mRowHeight) >= LBound(mRowHeight) Then
'    For J = LBound(mRowHeight) To UBound(mRowHeight)
'        iRno = iRno + 1
'        If Val(mRowHeight(J)) > 0 Then
'            With mWs.Rows(iRno)
'                .RowHeight = Val(mRowHeight(J))
'                .HorizontalAlignment = xlCenter
'                .VerticalAlignment = xlTop
'                .WrapText = True
'            End With
'        End If
'    Next
'End If
'''FreezeAt
'mWs.Range(p.FreezeAt).Select
'Set_Zoom mWs, 80
'''HideRows
'Dim mHideRows$(): mHideRows = Split(p.HideRows, CtComma)
'For J = LBound(mHideRows) To UBound(mHideRows)
'    mWs.Rows(mHideRows(J)).EntireRow.Hidden = True
'Next
'GoTo X
'R: ss.R
'E: Gen_Ws = True: ss.B cSub, cMod, ""
'X: Clr_Sts
'End Function
'Function Gen_Xls(pFxFm$, pFxTo$ _
'        , Optional pFb_DtaSrc$ = "" _
'        , Optional pLnWsRmv$ = "" _
'        , Optional pLExpr$ = "" _
'        , Optional pCollMacro As VBA.Collection _
'        , Optional pKeepWbOpnAndNotSav As Boolean = False _
'        , Optional oWb As Workbook _
'        ) As Boolean
'Const cSub$ = "Gen_Xls"
'Dim mChk_ChartTit As Boolean
'Dim mChk_CmdTxt As Boolean
'Dim mChk_Prm As Boolean
'mChk_ChartTit = False
'mChk_CmdTxt = False
'mChk_Prm = False
'
''Aim       : Generate {pFxTo} from the template Xls {pFxFm} by copying and refreshing using CurrentDb (or {pFb_DtaSrc} is given) as datasource
''Parameters:
'''RmvWsLst : If given, the ws list will be removed before refresh.
'''pKeepWbOpnAndNotSav : If True, the generated workbook will be kept open and not save, so that it can be further processed without re-open.
'''pLExpr       : If given, all CtCommandText of all Pc,Pt will add a where clause of where {Filter} to the end.
'''                Assume there is no where clause in the CtCommandText
'''pHdrMacromColl : If given, each worksheet's page header string will be scanned to do the macro substiution
''==Start==
'If mChk_Prm Then
'    Shw_Dbg cSub, cMod, "Check each of the param", "pFxFm,pFxTo,pLnWsRmv,pKeepWbOpnAndNotSav,pLExpr,pCollMacro", pFxFm, pFxTo, pLnWsRmv, pKeepWbOpnAndNotSav, pLExpr, ToStr_Coll(pCollMacro)
'    Stop
'End If
'
''Copy pFxFm to pFxTo and open in Xls by Calling <CopyAndOpen>
'StsShw Fmt_Str("GenXls 1. Create File: ToFil[{0}]...", pFxTo)
'If FxCpyAndOpn(oWb, pFxFm, pFxTo, True) Then ss.A 1: GoTo E
'
''Remove <RmvWsLst> if needed
'Dim iWs As Worksheet, iCht As Chart, V
'If pLnWsRmv <> "" Then
'    Dim AnWs$(): AnWs = Split(pLnWsRmv, CtComma)
'    gXls.DisplayAlerts = False
'    For Each V In AnWs
'        For Each iWs In oWb.Worksheets
'            If InStr(iWs.Name, Trim(V)) > 0 Then iWs.Delete
'        Next
'        For Each iCht In oWb.Charts
'            If InStr(iCht.Name, Trim(V)) > 0 Then iCht.Delete
'        Next
'    Next
'    gXls.DisplayAlerts = True
'End If
'
''Refresh PivotCache
'If Rfh_Wb(oWb, pLExpr, Fct.NonBlank(pFb_DtaSrc, CurrentDb.Name)) Then ss.A 2: GoTo E
'
''Page Header Macro Substiution
'If Not IsNull(pCollMacro) Then
'    ''Build mKey() & mVal() from pCollMacro
'    Dim AyK$(), AyV$(): Cv_CollKvStr_To2Ay pCollMacro, AyK, AyV
'
'    For Each iWs In oWb.Worksheets
'        SysCmd acSysCmdSetStatus, Fmt_Str("GenXls 2. Set Worksheet Heading: Ws[{0}]...", iWs.Name)
'        If Repl_WsPagSetup(iWs.PageSetup, AyK, AyV) Then ss.A 3: GoTo E
'        If Repl_WsChtObj(iWs, AyK, AyV) Then ss.A 4: GoTo E
'    Next
'    Dim iChart As Chart
'    For Each iChart In oWb.Charts
'        StsShw Fmt_Str("GenXls 3. Set Chart Heading: Chart[{0}]...", iChart.Name)
'        If Repl_WsChtTit(iChart.ChartTitle, AyK, AyV) Then ss.A 5: GoTo E
'    Next
'End If
''Hide all ws with name begin with "data"
'For Each iWs In oWb.Worksheets
'    If Left(iWs.Name, 4) = "data" Then iWs.Visible = xlSheetHidden
'Next
'
'Dim mVisible As Boolean
'If mChk_CmdTxt Then
'    Shw_Dbg cSub, cMod, "Check.CtCommandText of Pivot Table & Query Table", "pLExpr", pLExpr
'    mVisible = gXls.Visible
'    If Not gXls.Visible Then gXls.Visible = True
'    Lst_CmdTxt oWb, 0
'    Stop
'    gXls.Visible = mVisible
'End If
'Dim mMsg$
'If mChk_ChartTit Then
'    mMsg = "Check Generated Xls" & vbLf & vbLf & _
'    "Is {NmSess} {NmDta} of each chart replaced properly"
'
'    Shw_Dbg cSub, cMod, "Check.ChartTitle of Charts", "pCollMacro", ToStr_Coll(pCollMacro)
'    mVisible = gXls.Visible
'    If Not gXls.Visible Then gXls.Visible = True
'    Stop
'    gXls.Visible = mVisible
'End If
'Clr_Sts
'If WbFmtQt(oWb) Then ss.A 1, "Cannot format Wb": GoTo E
'If pKeepWbOpnAndNotSav Then Exit Function
'GoTo X
'R: ss.R
'E: Gen_Xls = True: ss.B cSub, cMod, "pFxFm,pFxTo,pFb_DtaSrc,pLnWsRmv,pLExpr,pCollMacro,pKeepWbOpnAndNotSav", pFxFm$, pFxTo$, pFb_DtaSrc$, pLnWsRmv$, pLExpr$, ToStr_Coll(pCollMacro), pKeepWbOpnAndNotSav
'X: Cls_Wb oWb, True
'    Clr_Sts
'End Function
'Function Gen_Xls__Tst()
'Const cSub$ = "Gen_Xls_Tst"
'Dim mCase%: mCase = 3
'Dim mFfnFm$, mFfnTo$, mWb As Workbook, mWs As Worksheet
'Select Case mCase
'Case 1
'    Dim mYYYYMMDD$: mYYYYMMDD = "2006_01_01"
'    mFfnFm = Sffn_Tp("TskLst")
'    mFfnTo = Sffn_Rpt("TskLst", mYYYYMMDD)
'    Gen_Xls_Tst = Gen_Xls(mFfnFm, mFfnTo)
'    If Opn_Wb_RW(mWb, mFfnTo, True) Then ss.A 2: GoTo E
'Case 2
'    'Aim: assume the TmpInqAR mdb is always created before ExpAR can be called
'    If Not Fct.Start("Export current AR inquiry data to Excel (c:\tmp\ARCollection\ARInq.xls)") Then Exit Function
'    mFfnFm = Sffn_Tp("ARInq")
'    mFfnTo = Sdir_TmpApp() & "ARInq.xls"
'    If Gen_Xls(mFfnFm, mFfnTo, Sffn_TmpAppUsrMdb()) Then ss.A 3: GoTo E
'    If Opn_Wb_RW(mWb, mFfnTo, True) Then ss.A 4: GoTo E
'    Set mWs = mWb.Sheets(1)
'    With mWs.Range("A3:AD3")
'        .HorizontalAlignment = xlLeft
'        .VerticalAlignment = xlTop
'        .WrapText = True
'        .MergeCells = True
'    End With
'    mWb.Save
'Case 3
'    mFfnTo = Sdir_Tmp & "a.xls"
'    mFfnFm = Sffn_Tp("CusLstForEdt")
'    If Gen_Xls(mFfnFm, mFfnTo, Sdir_Hom & "ARCollection.Mdb") Then ss.A 3: GoTo E
'    If Opn_Wb_RW(mWb, mFfnFm, True) Then ss.A 4: GoTo E
'    gXls.Visible = True
'End Select
'Exit Function
'R: ss.R
'E: Gen_Xls_Tst = True: ss.B cSub, cMod
'End Function
'
'

