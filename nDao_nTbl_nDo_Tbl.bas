Attribute VB_Name = "nDao_nTbl_nDo_Tbl"
Option Compare Database
Option Explicit

Sub TblAddFld(T, F$, Ty As DatabaseTypeEnum, Optional A As database)
If TblHasFld(T, F, A) Then Exit Sub
Dim B$: B = DaoTySqs(Ty)
Dim S$: S = FmtQQ("Alter Table [?] Add Column [?] ?", T, F, B)
DbRunSql S, A
End Sub

Sub TblAddFmTblByNm(ToTbl$, FmTbl$, ToFld$, Optional FmFld$ = "", Optional A As database)
'Aim: Add Distinct FmTbl!FmFld into ToTbl!ToFld for those not exist
Dim FmFld1$: FmFld1 = NonBlank(FmFld, ToFld)
Const Sql$ = "Insert into [{To}] ({ToFld}) select Distinct {FmFld} from [{Fm}]" & _
    " where {FmFld} not in (Select {ToFld} from [{To}])" & _
    " and {FmFld}<>''"
DbRunSqlNmAv Sql, ApAv(ToTbl, ToFld, FmFld1, FmTbl), A
End Sub

Sub TblAddTblByNm__Tst()
TblCrt_ByFldDclStr "#aa", "aa text 10"
TblCrt_ByFldDclStr "#bb", "bb text 10"
SqlRun "Insert into [#aa] values('1')"
SqlRun "Insert into [#aa] values('2')"
SqlRun "Insert into [#bb] values('2')"
SqlRun "Insert into [#bb] values('2')"
SqlRun "Insert into [#bb] values('3')"
SqlRun "Insert into [#bb] values('3')"
SqlRun "Insert into [#bb] values('3')"
TblAddFmTblByNm "#aa", "#bb", "aa", "bb"
DoCmd.OpenTable "#aa"
Stop
DoCmd.Close acTable, "#aa"
TblDrp "#aa"
TblDrp "#bb"
End Sub

Function TblApdFmTbl(TarTn$, pNmtSrc$, pNKFld As Byte, Optional pNKFldRmv = 0) As Boolean
'Aim: Add/Upd {TarTn} by {pNmtSrc}.  Both has same {pNKFld} of PK.  All fields in {pNmtSrc} should all be found in {TarNmt}
'     If pNKFldRmv>0 then some record in {TarTn} will be remove if they does not exist in pNmtSrc having first {pNKFldRmv} as the matching keys
'     Example, Tar & Src: a,b,c, x,y,z
'              pNKFld   : 3
'              pNKFld   : 2
'              Tar: 1,1,3, ..... Src: 1,1,4, ...
'                   1,1,4, .....      1,1,5 ...
'                   1,1,5, .....      1,1,6, ...
'                 : 1,2,3, .....
'                   1,2,4, .....
'                   1,2,5, .....
'    After
'              Tar: 1,1,4
'                   1,1,5
'                   1,1,6
Dim mSqlAdd$, mSqlUpd$, mSqlDlt$
'mSqlAdd, mSqlUpd, mSqlDlt,
'SqsOfAddUpdDlt TarTn, pNmtSrc, pNKFld, pNKFldRmv '
Stop
StsShw Fmt_Str("TblApdFmTbl: Adding [{0}] to [{1}] with pNKFld=[{2}] & pNKFldRmv=[{3}]", pNmtSrc, TarTn, pNKFld, pNKFldRmv)
Run_Sql mSqlAdd
Run_Sql mSqlUpd
If pNKFldRmv > 0 Then
    Run_Sql mSqlDlt
End If
Clr_Sts
End Function

Function TblApdFmTbl__Tst()
'Create cNmtTar & cNmtSrc
Const cNPK% = 2
Const cNKFldRmv% = 0
Const cNmtTar$ = "#TblApdFmTbl_Tar", cLoFldTar$ = "aa Int, bb Int, t1 Text 10, t2 Text 10, t3 Text 10, t4 Text 10"
Const cNmtSrc$ = "#TblApdFmTbl_Src", cLoFldSrc$ = "aa Int, bb Int, t1 Text 10, t2 Text 10, t3 Text 10"

TblCrt_ByFldDclStr cNmtTar, cLoFldTar, cNPK
TblCrt_ByFldDclStr cNmtSrc, cLoFldSrc, cNPK

'Do Add data to cNmtTar & cNmtSrc
Do
    Dim J%
    Const cNRecTar% = 3 + 1
    Dim mAyRec_Tar$(cNRecTar - 1)
    mAyRec_Tar(0) = "1,0,'aa0','bb0','cc0','dd0'"
    mAyRec_Tar(1) = "1,1,'aa1','bb1','cc1','dd2'"
    mAyRec_Tar(2) = "1,2,'aa2','bb2','cc2','dd2'"
    mAyRec_Tar(3) = "1,3,'aa3','bb3','cc3','dd3'"
'    mAyRec_Tar(4) = "1,4,'aa4','bb4','cc4','dd4'"
'    mAyRec_Tar(5) = "1,5,'aa5','bb5','cc5','dd5'"
'    mAyRec_Tar(6) = "1,6,'aa6','bb6','cc6','dd6'"
'    mAyRec_Tar(7) = "1,7,'aa7','bb7','cc7','dd7'"
    Const cNRecSrc% = 7
    Dim mAyRec_Src$(cNRecSrc - 1): J = 0
    mAyRec_Src(J) = "1,1,'AA1','BB1','CC1'": J = J + 1
    mAyRec_Src(J) = "1,2,'AA2','BB2','CC2'": J = J + 1
    mAyRec_Src(J) = "1,3,'AA3','BB3','CC3'": J = J + 1
    mAyRec_Src(J) = "1,4,'AA4','BB4','CC4'": J = J + 1
    mAyRec_Src(J) = "1,5,'AA5','BB5','CC5'": J = J + 1
    mAyRec_Src(J) = "1,6,'AA6','BB6','CC6'": J = J + 1
    mAyRec_Src(J) = "1,7,'AA7','BB7','CC7'": J = J + 1

    Dim mSql$
    For J% = 0 To cNRecTar - 1
        mSql = Fmt_Str("Insert into [{0}] values ({1})", cNmtTar, mAyRec_Tar(J))
        If Run_Sql(mSql) Then Stop
    Next
    For J% = 0 To cNRecSrc - 1
        mSql = Fmt_Str("Insert into [{0}] values ({1})", cNmtSrc, mAyRec_Src(J))
        If Run_Sql(mSql) Then Stop
    Next
Loop Until True

If TblApdFmTbl(cNmtTar, cNmtSrc, cNPK, cNKFldRmv) Then Stop
DoCmd.OpenTable cNmtTar
End Function

Sub TblBrw(T, Optional D As database)
DbAppa(D).DoCmd.OpenTable T, acViewNormal, acReadOnly
End Sub

Sub TblCrtIdx(T, IdxNm$, FnStr$, Optional IsUniq As Boolean = False, Optional A As database)
'Aim: Create {pIdx} on {T} by {FnStr}
TblDrpIdx T, IdxNm, A
Dim S$:
    Dim F$
    Dim Uniq$
    Uniq = IIf(IsUniq, "UNIQUE ", "")
    F = FnStrLvc(FnStr)
    S = FmtQQ("Create ?Index ? on ? (?)", Uniq, IdxNm, T, F)

DbRunSql S, A
End Sub

Function TblCrtIdx__Tst()
TblCrt_ByFldDclStr "aa", "aa text 10, bb text 10"
TblCrtIdx "aa", "U01", "aa,bb"
TblCrtIdx "aa", "U01", "bb,aa"
DoCmd.OpenTable "aa", acViewDesign
End Function

Sub TblDrp(T, Optional A As database)
Dim Db As database: Set Db = DbNz(A)
If DbHasTbl(T, Db) Then Db.Execute FmtQQ("Drop Table [?]", T)
End Sub

Sub TblDrpIdx(T, IdxNm$, Optional A As DAO.database)
If Not TblHasIdx(T, IdxNm, A) Then Exit Sub
Dim S$: S = FmtQQ("Drop Index [{0}] on [{1}]", IdxNm, T)
DbRunSql S, A
End Sub

Function TblHasPrp(T As TableDef, PrpNm$) As Boolean
TblHasPrp = PrpIsExist(PrpNm, T.Properties)
End Function

Sub TblIns(T, FnStr$, Av(), Optional A As database)
DbRunSql SqsOfIns(T, FnStr, Av), A
End Sub

Function TblInsRecBy2Id(pNmt$, pNmFld_Pk1$, pNmFld_Pk2$, pId1&, pId2&, FnStr$, ParamArray pAp()) As Boolean
'Aim: Add/Update a record to {pNmt} with {FnStr} & {pAyV}
'     Assume {pNmt} has  Id fields as Pk
Const cSub$ = "TblInsRecBy2Id"
Dim mRs As DAO.Recordset
Dim mSql$: mSql = Fmt_Str("Select * from {0} where {1}={2} and {3}={4}", pNmt, pNmFld_Pk1, pId1, pNmFld_Pk2, pId2)
If Opn_Rs(mRs, mSql) Then ss.A 1: GoTo E
With mRs
    If .AbsolutePosition = -1 Then
        .AddNew
        .Fields(pNmFld_Pk1).Value = pId1
        .Fields(pNmFld_Pk2).Value = pId2
    Else
        .Edit
    End If
    If Set_Rs_ByLpVv(mRs, FnStr, CVar(pAp)) Then ss.A 1: GoTo E
    .Update
End With
GoTo X
R:
E:
X:
RsCls mRs
End Function

Sub TblInsRecByUKey(oId&, pNmt$, pUKey_NmFld$, pUKey_Val)
'Aim: Add a record to {pNmt} with {pUKey_NmFld}, {pUKey_Val}
'     Assume first field is the Id & AutoField field and with be returned in {oId}
Const cSub$ = "TblInsRecByUKey"
Dim mRs As DAO.Recordset
Dim mSql$: mSql = Fmt_Str("Select * from {0} where {1}='{2}'", pNmt, pUKey_NmFld, pUKey_Val)
If Opn_Rs(mRs, mSql) Then ss.A 1: GoTo E
With mRs
    If .AbsolutePosition = -1 Then
        .AddNew
        oId = .Fields(0).Value
        .Fields(pUKey_NmFld).Value = pUKey_Val
        .Update
    Else
        oId = .Fields(0).Value
    End If
    .Close
End With
R:
E:
End Sub

Function TblInsRecByUKey__Tst()
Dim Mid&
TblInsRecByUKey Mid, "xx", "bb", "1234"
End Function

Sub TblInsRecByUKey_n_LpAp(oId&, pNmt$, pUKey_NmFld$, pUKey_Val$, FnStr$, ParamArray pAp())
Const cSub$ = "TblInsRecUKey_n_LpAp"
'Aim: Add or Update a record to {pNmt} with {pUKey_NmFld}, {pUKey_Val} by {FnStr} and {mAyV} & return {oId}
'     Assume {pNmt} has an unique key of a string field {pUKey_NmFld}
'     Assume first field is the Id & AutoField field and with be returned in {oId}
Dim mRs As DAO.Recordset
Dim mSql$: mSql = Fmt_Str("Select * from {0} where {1}='{2}'", pNmt, pUKey_NmFld, pUKey_Val)
If Opn_Rs(mRs, mSql) Then ss.A 1: GoTo E
With mRs
    If .AbsolutePosition = -1 Then
        .AddNew
        oId = .Fields(0).Value
        .Fields(pUKey_NmFld).Value = pUKey_Val
        If Set_Rs_ByLpVv(mRs, FnStr, CVar(pAp)) Then ss.A 1: GoTo E
        .Update
    Else
        .Edit
        oId = .Fields(0).Value
        If Set_Rs_ByLpVv(mRs, FnStr, CVar(pAp)) Then ss.A 2: GoTo E
        .Update
    End If
    .Close
End With
Exit Sub
R: ss.R
E:
End Sub

Function TblJnRec(pNmtFm$, pNmtTo$, pLoKey$, pNmFld_NRec$, Optional pSepChr$ = CtCommaSpc) As Boolean
'Aim: Create {pNmtTo} from {pNmtFm}.  The new table will have Key fields as list in {pLoKey} plus 1 more field [Lst{pNmFld}].
'     The value of this field [Lst{pNmFld}] is coming those records in {pNmtFm} of the current key.
Const cSub$ = "TblJnRec"
'Build Empty {pNmtTo}
If Dlt_Tbl(pNmtTo) Then ss.A 1: GoTo E
Dim mSql$: mSql = Fmt_Str("Select Distinct {0} into {1} from {2} where False", pLoKey, pNmtTo, pNmtFm)
If Run_Sql(mSql) Then ss.A 1: GoTo E
mSql = Fmt_Str("Alter table {0} Add COLUMN Lst{1} Memo", pNmtTo, pNmFld_NRec)
If Run_Sql(mSql) Then ss.A 2: GoTo E
'Loop {pNmtFmt} having break @ mAnFldKey()
Dim mAnFldKey$(): mAnFldKey = Split(pLoKey, CtComma)
Dim NKey%: NKey = Sz(mAnFldKey)
Dim mRs As DAO.Recordset: Set mRs = CurrentDb.OpenRecordset(Fmt_Str("Select {0},{1} from {2} order by {0},{1}", pLoKey, pNmFld_NRec, pNmtFm))
With mRs
    If .AbsolutePosition = -1 Then .Close: Exit Function
    ReDim mAyLasKeyVal(NKey - 1)
    Dim mLasKeyVal$:
    Dim J%: For J = 0 To NKey - 1
        mAyLasKeyVal(J) = .Fields(mAnFldKey(J)).Value
    Next
    If Fnd_LvFmRs(mLasKeyVal, mRs, pLoKey) Then ss.A 3: GoTo E

    Dim mLst$
    Dim mSql_Tp$: mSql_Tp = Fmt_Str("Insert into {0} ({1},Lst{2}) values ", pNmtTo, pLoKey, pNmFld_NRec) & "({0},'{1}')"
    While Not .EOF
        'If !Dte = #2/22/2007# And !Txt = "1010" And !InstId = 10 Then Stop
        If IsSamKey_ByAnFldKey(mRs, mAnFldKey, mAyLasKeyVal) Then
            mLst = Add_Str(mLst, .Fields(pNmFld_NRec).Value, pSepChr)
        Else
            '
            mSql = Fmt_Str(mSql_Tp, mLasKeyVal, mLst)
            If Run_Sql(mSql) Then ss.A 4: GoTo E

            For J = 0 To NKey - 1
                mAyLasKeyVal(J) = .Fields(mAnFldKey(J)).Value
            Next
            If Fnd_LvFmRs(mLasKeyVal, mRs, pLoKey) Then ss.A 3: GoTo E
            mLst = mRs.Fields(pNmFld_NRec).Value
        End If
        .MoveNext
    Wend
    mSql = Fmt_Str(mSql_Tp, mLasKeyVal, mLst)
    If Run_Sql(mSql) Then ss.A 5: GoTo E
    .Close
End With
Exit Function
R: ss.R
E:
End Function

Function TblJnRec__Tst()
Const cSub$ = "TblJnRec_Tst"
Dim mNmtFm$: mNmtFm = "tmpBldTbl_NRec2Lst_Fm"
Dim mNmtTo$: mNmtTo = "tmpBldTbl_NRec2Lst_To"
Dim mLoKey$: mLoKey = "InstId,Dte,Txt"
Dim mNmFld_NRec$: mNmFld_NRec = "Num"
Dim mSepChr$: mSepChr = CtComma

Dim mBldTstTbl As Boolean: mBldTstTbl = False
If mBldTstTbl Then
    If Dlt_Tbl(mNmtFm) Then ss.A 1: GoTo E
    If Run_Sql(Fmt_Str("Create table {0} (InstId Long, Dte Date, Txt Text(10), Num Long)", mNmtFm)) Then ss.A 1: GoTo E
    Dim iInstId%: For iInstId = 0 To 10
        Dim iDte%: For iDte = 0 To 10
            Dim iTxt%: For iTxt = 1000 To 1010
                Dim iNum%: For iNum = 2000 To 2010
                    If Run_Sql(Fmt_Str("insert into {0} (InstId, Dte, Txt, Num) values ({1}, #{2}#, '{3}', {4})", mNmtFm, iInstId, Date + iDte, iTxt, iNum)) Then ss.A 1: GoTo E
                Next
            Next
        Next
    Next
End If
If TblJnRec(mNmtFm, mNmtTo, mLoKey, mNmFld_NRec, mSepChr) Then ss.A 3: GoTo E
Exit Function
R: ss.R
E:
End Function

Function TblPutCell(Qry_or_Tbl_Nm$, Rg As Range _
    , Optional SrcFb$ = "" _
    , Optional pNoExpTim As Boolean = False _
    ) As Boolean
'Aim: Read data from table {SrcFb}!{QryNmt} to QryNmt.Destination
Const cSub$ = "TblPutCell"
On Error GoTo R
Dim mWs As Worksheet: Set mWs = Rg.Parent

'Build Qt
Clr_Qt mWs
Dim mFbSrc$
If SrcFb = "" Then
    mFbSrc = CurrentDb.Name
Else
    mFbSrc = SrcFb
End If
Dim mQt As QueryTable: Set mQt = mWs.QueryTables.Add(CnnStr_Mdb(mFbSrc), Rg)
With mQt
    Dim mSql$: If BldSql_Qt(mSql, Rg, Qry_or_Tbl_Nm, SrcFb) Then ss.A 2: GoTo E
    .CtCommandType = xlCmdSql
    .CtCommandText = mSql
    .BackgroundQuery = False
    .AdjustColumnWidth = False
    .FillAdjacentFormulas = False
    .MaintainConnection = False
    .PreserveColumnInfo = True
    .PreserveFormatting = True
    .FieldNames = False
End With

'Fill Data
Shw_AllDta mWs
Dim mAdr$: mAdr = Rg.Address
mWs.Range(mWs.Cells(Rg.Row, 1), mWs.Cells(65536, 1)).EntireRow.ClearFormats
mWs.Range(mWs.Cells(Rg.Row, 1), mWs.Cells(65536, 1)).EntireRow.Delete
Set Rg = mWs.Range(mAdr)

If Rfh_Qt(mQt) Then ss.A 3: GoTo E
Dim mNRec&: mNRec = mQt.ResultRange.Rows.Count

'Fmt Qt
If WsFmt(Rg, mNRec, 3) Then ss.A 4: GoTo E

Clr_Qt mWs
Exit Function
R: ss.R
E:
End Function

Sub TblRbr(pTbl$, pNmfld_ToRbr$, Optional pStart As Byte = 1, Optional pStp As Byte = 1)
Dim mRs As DAO.Recordset: Set mRs = CurrentDb.OpenRecordset("Select {0} from {1} order by {0}", pNmfld_ToRbr, pTbl)
RsRbr mRs, pStart, pStp
End Sub

Function TblRen(pNmtFm$, pNmtTo$) As Boolean
Const cSub$ = "Ren_Tbl_ByNmt"
If Dlt_Tbl(pNmtTo) Then ss.A 1, "pNmtTo cannot be deleted": GoTo E
On Error GoTo R
CurrentDb.TableDefs(pNmtFm).Name = pNmtTo
Exit Function
R: ss.R
E:
End Function

Function TblRenPfx(FmPfx$, ToPfx$) As Boolean
Dim L%: L = Len(FmPfx)
Dim iTbl As TableDef: For Each iTbl In CurrentDb.TableDefs
    If Left(iTbl.Name, L) = FmPfx Then
        Debug.Print "Renaming ... "; iTbl.Name
        iTbl.Name = ToPfx & Mid$(iTbl.Name, L + 1)
    End If
Next
End Function
Sub TblCpy(Src, Tar, FldDic As Dictionary, A As database)
Dim Db As database: Set Db = DbNz(A)
TblAsstExist Src, Db
TblDrpEns Tar, Db
Dim Ay$(): Ay = DicSy(FldDic, "{K} as {V}")
Dim Sel$: Sel = JnComma(Ay)
Const C$ = "Select ? into [?] from [?]"
Dim S$: S = FmtQQ(C, Sel, Tar, Src)
SqlRun S, Db
End Sub

Sub TblDrpEns(T, Optional A As database)
If TblIsExist(T, A) Then TblDrp T, A
End Sub

Function TblRenToBackup(ToPfx$) As Boolean
'Aim: Rename all linked table by adding {ToPfx}
Const cSub$ = "TblRenToBackup"
On Error GoTo R
If ToPfx = "" Then ss.A 1, "ToPfx cannot be blank": GoTo E
Dim iTbl As TableDef
For Each iTbl In CurrentDb.TableDefs
    If iTbl.Connect <> "" Then iTbl.Name = ToPfx & iTbl.Name
Next
Exit Function
R: ss.R
E:
End Function

Function TblToFxmll__Tst()
TblWrtFxml "permit", "c:\tmp\mstBrand.xml"
End Function

Sub TblWrtFb(TnyOpt, TarFb$, Optional pCrtIfNotExist As Boolean = False, Optional A As database)

'Aim: Currentdb db's {pLikNmt} tables to {TarFb}
Const cSub$ = "Snd_Tbl_ToMdb"
If VBA.Dir(TarFb) = "" Then
    If Not pCrtIfNotExist Then ss.A 1, "Given {TarFb} not exist": GoTo E
    FbNew TarFb
End If
Dim mAnt$(): ' If Fnd_Ant_ByLik(mAnt, pLikNmt) Then ss.A 4: GoTo E
Dim J%
For J = 0 To Sz(mAnt) - 1
    Dim mSql$: mSql = Fmt_Str("Select * into {0} in '{1}' from {0}", mAnt(J), TarFb)
    If Run_Sql(mSql) Then ss.A 5: GoTo E
Next
Exit Sub
R: ss.R
E:
End Sub

Sub TblWrtFxml(T, Fxml$)
Application.ExportXML acExportTable, T, Fxml, , , , , acEmbedSchema
End Sub

