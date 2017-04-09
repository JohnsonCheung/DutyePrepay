Attribute VB_Name = "ZZ_xRfh"

'Option Compare Text
'Option Explicit
'Option Base 0
'Const cMod$ = cLib & ".xRfh"
'Function Rfh_InUse(Optional pNmt$ = "tblLnkTbl") As Boolean
''Aim: Refresh tblLnkTbl->InUse from all queries to see if tblLnkTbl->Nmt is used by any queries in CurrentDb
''     Assume {pNmt}: NmTbl, InUse
'Const cSub$ = "Rfh_InUse"
'On Error GoTo R
'Dim mNmt$: mNmt = Q_SqBkt(pNmt)
'If Run_Sql("Update " & mNmt & " set InUse=False") Then ss.A 1: GoTo E
'Dim mRs As DAO.Recordset: If Opn_Rs(mRs, "Select InUse,NmTbl from " & mNmt$) Then ss.A 2: GoTo E
'With mRs
'    While Not .EOF
'        If IsStrExistInQry(CStr(!NmTbl.Value)) Then
'            .Edit
'            !InUse = True
'            .Update
'        End If
'        .MoveNext
'    Wend
'End With
'GoTo X
'R: ss.R
'E: Rfh_InUse = True: ss.B cSub, cMod
'X: RsCls mRs
'End Function
'Function Rfh_Lnk_ByRsLnkDef(pRsLnkDef As DAO.Recordset) As Boolean
'Const cSub$ = "Rfh_Lnk_ByRsLnkDef"
''Aim: Delete all link tables in CurrentDb
''     Create NonBlank(!NmtNew, !Nmt) in currentdb to link !Nmt in !Ffn of Type !NmLnkTyp
''     Assume pRsLnkDef has structure: Nmt, InFfn, NmLnkTyp, NmtNew
'If Dlt_Tbl_ByLnk Then ss.A 1: GoTo E
'On Error GoTo R
'With pRsLnkDef
'    While Not .EOF
'        Select Case !NmTypLnk
'        Case "XlsWs"
'            If TblCrt_FmLnkWs(!InFfn, !Nmt, Nz(!NmtNew, "")) Then ss.A 2: GoTo E
'        Case "MdbTbl"
'            If TblCrt_FmLnkNmt(!InFfn, !Nmt, Nz(!NmtNew, "")) Then ss.A 3: GoTo E
'        Case Else
'        'Case "TxtFil"
'            ss.A 4, "Unexpected Link Type", , "!Nmt,InFfn,!NmLnkTyp,!NmtNew", !Nmt, !InFfn, !NmLnkTyp, !NmtNew: GoTo E
'        End Select
'        .MoveNext
'    Wend
'End With
'Exit Function
'R: ss.R
'E: Rfh_Lnk_ByRsLnkDef = True: ss.B cSub, cMod, "pRsLnkDef", ToStr_Rs(pRsLnkDef)
'End Function
'Function Rfh_LnkV1(pNmLgc$ _
'    , Optional pLn$ _
'    , Optional pV0$ = "" _
'    , Optional pV1$ = "" _
'    , Optional pV2$ = "" _
'    , Optional pV3$ = "" _
'    , Optional pV4$ = "" _
'    , Optional pV5$ = "" _
'    , Optional pV6$ = "" _
'    , Optional pV7$ = "" _
'    , Optional pV8$ = "" _
'    , Optional pV9$ = "" _
'    , Optional pV10$ = "" _
'    , Optional pV11$ = "" _
'    , Optional pV12$ = "" _
'    , Optional pV13$ = "" _
'    , Optional pV14$ = "" _
'    , Optional pV15$ = "") As Boolean
''Aim:   Delete all linked tables in currentdb
''       relink all those link table described in [tblLnkTblV1] of {NmLgc}
''       [tblLnkTblV1]=NmLgc,Nmt,InFfn,LnkNmt,NmNew,TypLnk
'Const cSub$ = "Rfh_LnkV1"
'On Error GoTo R
'Dim mSql$: mSql = "SELECT InFfn, FfnMacro FROM tblLnkTblV1 where Trim(Nz(FfnMacro,''))<>'' and NmLgc='" & pNmLgc & CtSngQ
'With CurrentDb.OpenRecordset(mSql)
'    While Not .EOF
'        .Edit
'        !InFfn.Value = Fmt_Str_ByLpAp(CStr(!FfnMacro.Value), pLn, pV0, pV1, pV2, pV3, pV4, pV5, pV6, pV7, pV8, pV9, pV10, pV11, pV12, pV13, pV14, pV15)
'        .Update
'        .MoveNext
'    Wend
'    .Close
'End With
'
'mSql = "SELECT Nmt, InFfn, NmtNew, NmTypLnk" & _
'" FROM tblLnkTblV1 lt INNER JOIN tblLnkTblV1Typ ltt ON lt.TypLnk = ltt.TypLnk" & _
'" where NmLgc='" & pNmLgc & CtSngQ
'Dim mRs As DAO.Recordset: Set mRs = CurrentDb.OpenRecordset(mSql)
'If Rfh_Lnk_ByRsLnkDef(mRs) Then ss.A 2: GoTo E
'mRs.Close
'Exit Function
'R: ss.R
'E: Rfh_LnkV1 = True: ss.B cSub, cMod, "pNmLgc", pNmLgc
'End Function

'Function Rfh_LnkV1__Tst()
'If Rfh_LnkV1("AddEle", "FbMeta", "p:\workingdir\Meta_Data.mdb") Then Stop
'End Function

'Function Rfh_Lnk(pTrc&) As Boolean
''Aim: Create link tables in CurDb for each record in "tblLnkTbl" & "tblLnkTblMdbSrc"
'Const cSub$ = "Rfh_Lnk"
'On Error GoTo R
'Dim mDirSess$: mDirSess = Fct.CurMdbDir & Format(pTrc, "00000000") & "\": If VBA.Dir(mDirSess, vbDirectory) = "" Then ss.A 1, , "[Sess Sub Dir] does not exist in currentDb", "CurDb", CurrentDb.Name: GoTo E
'Rfh_Lnk_Chk_tblLnkTbl
'Dim mFb_modU$:  mFb_modU = Sdir_PgmObj & "mda"
'Dim mFb_Dta$:   mFb_Dta = Sdir_Wrk & Fct.CurMdbNam & "_Data.mdb"
'Dim xFfn$, xNmtSrc$, mSql$, mLnkLib
'mSql = _
'"Select      Nmt,LnkLib,FbSrc" & _
'" from       tblLnkTbl_NewVer l" & _
'" inner join tblLnkTblMdbSrc  s" & _
'" on         l.MdbSrcId=s.MdbSrcId" & _
'" order by   LnkLib"
'With CurrentDb.OpenRecordset(mSql)
'    While Not .EOF
'        mLnkLib = Nz(!LnkLib, "")
'        xNmtSrc = !Nmt
'        If mLnkLib = "modU" Then
'            xFfn = mFb_modU
'        ElseIf mLnkLib = "" Then
'            xFfn = mFb_Dta
'        ElseIf Left(mLnkLib, 3) = "Tp:" Then
'            'Nmt    LnkLib
'            'aaa    Tp:TpNam!ssss
'            Dim mA$: mA = mID$(mLnkLib, 4)  'TpNam!ssss
'            Dim mP%: mP = InStr(mA, "!")    'Pos of !
'            Dim mTp$
'            If mP > 0 Then
'                mTp$ = Left(mA, mP - 1)     'TpNam
'                xNmtSrc = mID(mA, mP + 1)   'sss
'            Else
'                mTp$ = mA                   'TpNam
'            End If
'            If Fnd_Fn_By_Tp_n_CurFnn(mA, mTp, Fct.CurMdbNam) Then ss.A 1: GoTo E
'            xFfn = mDirSess & mA
'        ElseIf mLnkLib = "MdbSrc" Then
'            xFfn = !FbSrc
'        Else
'            xFfn = Sdir_PgmObj & mLnkLib
'        End If
'        'StsShw "Linking [" & !Nmt & "] to [" & xFfn & "] ........"
'        If TblCrt_FmLnkNmt(xFfn, xNmtSrc$, !Nmt) Then ss.A 1: GoTo E
'        .MoveNext
'    Wend
'    .Close
'End With
'GoTo X
'R: ss.R
'E: Rfh_Lnk = True: ss.B cSub, cMod, "pTrc", pTrc
'X: Clr_Sts
'End Function

'Function Rfh_Lnk__Tst()
'If Crt_SessDta(1) Then Stop
'If Rfh_Lnk(1) Then Stop
'End Function

'Private Function Rfh_Lnk_Chk_tblLnkTbl() As Boolean
''Aim: Check tblLnkTbl is in valid format
'Const cSub$ = "Rfh_Lnk_TblCrtLnkTbl"
'On Error GoTo R
'If Not Chk_Struct_Tbl("tblLnkTbl_NewVer", "Nmt,LnkLib,InUse,MdbSrcId") Then Exit Function
'If Run_Sql("Create table tblLnkTbl_NewVer (Nmt Text(50), LnkLib Text(50), InUse YesNo, MdbSrcId Integer)") Then ss.A 1: GoTo E
'Exit Function
'R: ss.R
'E: Rfh_Lnk_Chk_tblLnkTbl = True: ss.B cSub, cMod
'End Function
'Private Function Rfh_Lnk_Chk_tblLnkMdbSrc() As Boolean
''Aim: Check tblLnkTbl is in valid format
'Rfh_Lnk_Chk_tblLnkMdbSrc = Chk_Struct_Tbl("tblLnkMdbSrc", "Nmt,LnkLib,InUse,MdbSrcId")
'End Function
'Function Rfh_Pc(pPc As PivotCache) As Boolean
'Const cSub$ = "Rfh_Pc"
'On Error GoTo R
'pPc.Refresh
'Exit Function
'R: ss.R
'E: Rfh_Pc = True: ss.B cSub, cMod
'End Function
'Function Rfh_Pt(pPt As PivotTable) As Boolean
'Const cSub$ = "Rfh_Pt"
'On Error GoTo R
'pPt.RefreshTable
'Exit Function
'R: ss.R
'E: Rfh_Pt = True: ss.B cSub, cMod
'End Function

'Function Rfh_Wb(pWb As Workbook, Optional pLExpr$ = "", Optional pFb_DtaSrc$ = "") As Boolean
'Const cSub$ = "Rfh_Wb"
''Aim: Use current mdb as source to refresh given {pWorkbooks} data.
'Dim cMsg$: cMsg = "RfhWb Wb[" & pWb.Name & "] is refreshing {0} ..."
'Dim mCnnStr$: mCnnStr = CnnStr_Mdb(Fct.NonBlank(pFb_DtaSrc, CurrentDb.Name))
'With pWb
'    Dim iWs As Worksheet
'    'Refresh all [Listobject] in all worksheets of each workbook in <pWorkBooks>
'    StsShw Fmt_Str(cMsg, "ListObjects")
'    For Each iWs In .Worksheets
'        ''ListObject
'        Dim iLo As Excel.ListObject
'        Dim iQt As Excel.QueryTable
'        For Each iLo In iWs.ListObjects
'            Set iQt = iLo.QueryTable
'            With iQt
'                .Connection = mCnnStr
'                If pLExpr <> "" Then
'                    If .CtCommandType <> xlCmdSql Then ss.A 4, "Given CtCommand Type must be Sql": GoTo E
'                    If InStr(.CtCommandText, "where") > 0 Then ss.A 5, "Given Sql should have have where": GoTo E
'                    .CtCommandText = .CtCommandText & " WHERE " & pLExpr
'                End If
'                Rfh_Qt iQt
'            End With
'        Next
'    Next
'
'    'Refresh all [QueryTable] in all worksheets of each workbook in <pWorkBooks>
'    StsShw Fmt_Str(cMsg, "QueryTables")
'    For Each iWs In .Worksheets
'        ''QueryTable
'        For Each iQt In iWs.QueryTables
'            With iQt
'                .Connection = mCnnStr
'                If pLExpr <> "" Then
'                    If .CtCommandType <> xlCmdSql Then ss.A 4, "Given CtCommand Type must be Sql": GoTo E
'                    If InStr(.CtCommandText, "where") > 0 Then ss.A 5, "Given Sql should have have where": GoTo E
'                    .CtCommandText = .CtCommandText & " WHERE " & pLExpr
'                End If
'                Rfh_Qt iQt
'            End With
'        Next
'    Next
'
'    'Refresh all [PivotCache] of SourceTyp<>External in all workbooks as given in <pWorkbooks>
'    StsShw Fmt_Str(cMsg, "PivotCaches")
'    Dim iPc As Excel.PivotCache
'    For Each iPc In .PivotCaches
'        With iPc
'            If .SourceType <> xlDatabase Then
'                .Connection = mCnnStr
'                .BackgroundQuery = False
'                .MissingItemsLimit = xlMissingItemsNone
'                If pLExpr <> "" Then
'                    If .CtCommandType <> xlCmdSql Then ss.A 1, "Given CtCommand Type must be Sql": GoTo E
'                    If InStr(.CtCommandText, "where") > 0 Then ss.A 2, "Given Sql should have have where": GoTo E
'                    .CtCommandText = .CtCommandText & Cv_Str(pLExpr, " where ")
'                End If
'            End If
'            iPc.MissingItemsLimit = xlMissingItemsNone
'            Rfh_Pc iPc
'        End With
'    Next
'
'    'Refresh all [PivotTable] in all worksheets of each workbook in <pWorkBooks>
'    StsShw Fmt_Str(cMsg, "PivotTables")
'    For Each iWs In .Worksheets
'        Dim iPt As Excel.PivotTable
'        For Each iPt In iWs.PivotTables
'            Rfh_Pt iPt
'        Next
'    Next
'
'    ''ChartObj
''    StsShw Fmt_Str(cMsg, "Charts in Worksheet")
''    For Each iWs In .Worksheets
''        Dim iChartObj As ChartObject
''        For Each iChartObj In iWs.ChartObjects
''            If Not IsNothing(iChartObj.Chart.PivotLayout) Then
''                If Rfh_Pt(iChartObj.Chart.PivotLayout.PivotTable) Then ss.A 8: GoTo E
''            End If
''        Next
''    Next
''
''    'Refresh all [Charts]
''    StsShw Fmt_Str(cMsg, "Charts in Workbook")
''    Dim iChart As Chart
''    For Each iChart In pWb.Charts
''        If Not IsNothing(iChart.PivotLayout) Then
''            If Rfh_Pt(iChart.PivotLayout.PivotTable) Then ss.A 9: GoTo E
''        End If
''    Next
'End With
'GoTo X
'R: ss.R
'E: Rfh_Wb = True: ss.B cSub, cMod, "pWb,pLExpr,pFb_DtaSrc", ToStr_Wb(pWb), pLExpr, pFb_DtaSrc
'X: Clr_Sts
'End Function
'
'
'

