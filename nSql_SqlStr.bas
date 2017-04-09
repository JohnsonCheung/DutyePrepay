Attribute VB_Name = "nSql_SqlStr"
Option Compare Database
Option Explicit
Type SqlKw
    KwTy  As Byte
    KwLen As Byte
End Type

Sub SqlStr_AddUpdDtl(Tar$, Src$, NKeyFld As Byte, Optional NKeyFldRmv = 0)

'Aim: Add/Upd {Tar} by {Src}.  Both has same {NKeyFld} of PK.  All fields in {Src} should all be found in {Tar}
'     If NKeyFldRmv>0 then some record in {Tar} will be remove if they does not exist in Src having first {NKeyFldRmv} as the matching keys
'     Example, Tar & Src: a,b,c, x,y,z
'              NKeyFld   : 3
'              NKeyFld   : 2
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
Const cSub$ = "SqlStr_AddUpdDtl"
Dim mNmtTar$: mNmtTar = Rmv_SqBkt(Tar)
Dim mNmtSrc$: mNmtSrc = Rmv_SqBkt(Src)
Dim oSqlAdd$, oSqlUpd$, oSqlDlt$
Dim mJoin$, mSet$
Dim mAnKey$()
Do
    Dim mLnFld$: mLnFld = ToStr_Nmt(Src)
    Dim mAnFld$(): mAnFld = Split(mLnFld, CtComma)
    Dim mAnSet$(): If Ay_Cut(mAnKey, mAnSet, mAnFld, CInt(NKeyFld)) Then Er ""
    mJoin = StrExpand("t.{N}=s.{N}", mAnKey, , " and ")
    mSet = StrExpand("t.{N}=s.{N}", mAnSet)
Loop Until True

Dim mSql$
oSqlUpd = Fmt_Str("Update [{0}] t inner join [{1}] s on {2} set {3} ", mNmtTar, mNmtSrc, mJoin, mSet)
oSqlAdd = Fmt_Str("Insert into [{0}] Select s.* from [{1}] s left join [{0}] t on {2} where IsNull(t.{3})", mNmtTar, mNmtSrc, mJoin, mAnKey(0))
oSqlDlt = ""
'Find & Run {mSql} for Delete
If NKeyFldRmv > 0 Then
    Dim mExprPK$: mExprPK$ = Join(mAnKey, " & ")

    ReDim Preserve mAnKey(NKeyFldRmv - 1)
    Dim mExpr1stNFld$: mExpr1stNFld = Join(mAnKey, " & ")

    'DELETE *
    'From [$MdbS]
    'WHERE Mdb In (Select Mdb from [#MdbS])
    ' AND Mdb & Schm Not In (Select Mdb & Schm from [#MdbS])
    oSqlDlt = Fmt_Str("delete *" & _
        " from [{0}]" & _
        " where {2} in (Select {2} from [{1}])" & _
        " and {3} not in (Select {3} from [{1}])" _
        , mNmtTar, mNmtSrc, mExpr1stNFld$, mExprPK$)
End If
End Sub

Function SqlStrGpBy$(pGpBy$)
If pGpBy = "" Then Exit Function
SqlStrGpBy = " Group by " & pGpBy
End Function

Function SqlStrKwLen(S$) As SqlKw
Dim P%, X%, OTy%, OL%
P = InStrRev(S, "FROM "):                                          OTy = 1: OL = 5
X = InStrRev(S, "INNER JOIN "): If X > 0 Then If X > P Then P = X: OTy = 1: OL = 11
X = InStrRev(S, "LEFT JOIN "): If X > 0 Then If X > P Then P = X:  OTy = 1: OL = 10
X = InStrRev(S, "RIGHT JOIN "): If X > 0 Then If X > P Then P = X: OTy = 1: OL = 11

X = InStrRev(S, "INTO "): If X > 0 Then If X > P Then P = X:       OTy = 2: OL = 5
X = InStrRev(S, "DELETE "): If X > 0 Then If X > P Then P = X:     OTy = 2: OL = 7

X = InStrRev(S, "SELECT "): If X > 0 Then If X > P Then P = X:     OTy = 9: OL = 7
X = InStrRev(S, "AND "): If X > 0 Then If X > P Then P = X:        OTy = 9: OL = 4
X = InStrRev(S, "SET "): If X > 0 Then If X > P Then P = X:        OTy = 9: OL = 4
X = InStrRev(S, "ON "): If X > 0 Then If X > P Then P = X:         OTy = 9: OL = 3
X = InStrRev(S, "AS "): If X > 0 Then If X > P Then P = X:         OTy = 9: OL = 3
X = InStrRev(S, "WHERE "): If X > 0 Then If X > P Then P = X:      OTy = 9: OL = 9
X = InStrRev(S, "ORDER BY "): If X > 0 Then If X > P Then P = X:   OTy = 9: OL = 9
X = InStrRev(S, "GROUP BY "): If X > 0 Then If X > P Then P = X:   OTy = 9: OL = 9
SqlStrKwLen.KwLen = OL
SqlStrKwLen.KwTy = OTy
End Function

Function SqlStrOfCrt$(T, SqlFldList$)
Const C = "Create Table [?] (?)"
SqlStrOfCrt = FmtQQ(C, T, SqlFldList)
End Function

Function SqlStrOfIns$(T, FnStr$, Av())
Const C = "Insert into [?] (?) values (?)"
Dim F$: F = Join(FnStrBrk(FnStr), ",")
Dim U%: U = UB(Av)
Dim VV$(): ReDim VV(U)
Dim J%
For J = 0 To U
    VV(J) = SqlStrQuoteVar(Av(J))
Next
Dim V$: V = Join(VV, ",")
SqlStrOfIns = FmtQQ(C, T, F, V)
End Function

Function SqlStrOfSel$(T, Optional Sel = "*", Optional Where$, Optional OrdBy$)
SqlStrOfSel = FmtQQ("Select ? from [?] X??", Sel, T, SqlStrWhere(Where), SqlStrOrdBy(OrdBy))
End Function

Function SqlStrOfSel_CurHostDta$(oSql$, pRsUlSrc As DAO.Recordset, pNmtHost$, Optional oLExpr$)
'Aim: Build {oSql} to get data from Host by referring {pRsUlSrc}
'{pRsUlSrc} fmt: First N field is PK, Then [Changed], Then pair [xxx], [New xxx]
Const cSub$ = "SqlStrOfSel_CurHostDta"
Dim mSel$
oLExpr = ""
Dim J%: For J = 0 To pRsUlSrc.Fields.Count - 1
    If pRsUlSrc.Fields(J).Name = "Changed" Then
        Dim I%: For I = J + 1 To pRsUlSrc.Fields.Count - 1 - 5 Step 2 'Skip 5 columns at end
            mSel = Add_Str(mSel, pRsUlSrc.Fields(I).Name)
        Next
        Exit For
    End If
    With pRsUlSrc.Fields(J)
        Dim mA$: If Join_NmV(mA, .Name, .Value) Then ss.A 1: GoTo E
    End With
    oLExpr = Add_Str(oLExpr, mA, " and ")
Nxt:
Next
If mSel = "" Then ss.A 1, "mSel should be blank": GoTo E
If oLExpr = "" Then ss.A 2, "oLExpr should be blank": GoTo E
oSql = Fmt_Str("Select {0} from {1} where {2}", mSel, pNmtHost, oLExpr)
Exit Function
R: ss.R
E:
End Function

Function SqlStrOfSel1$(pSel$, pFm$, Optional pWhere$ = "", Optional pOrdBy$ = "")
Const cSqlSel$ = "Select {0} from {1}{2}{3}"
SqlStrOfSel1 = Fmt_Str(cSqlSel, pSel, Q_S(pFm, "[]"), SqlStrWhere(pWhere), SqlStrOrdBy(pOrdBy))
End Function

Function SqlStrOfUpd$(T, SetList$, Optional Where$)
SqlStrOfUpd = FmtQQ("Update [?] Set ??", T, SetList, SqlStrWhere(Where))
End Function

Function SqlStrOfUpd_ByRs(oSqlUpd$, pRs As DAO.Recordset, TarTn$, pLmPk$, Optional FnStr$ = "") As Boolean
''Aim: Build a {oSqlUpd} to Update table {TarTn} by the context in current record in {pRs}.
'      If {FnStr} is given, only those fields in the list will be Updated.
'      If {FnStr} is not given, all fields in {pRs} will be Updated
'Const cSub$ = "SqlStrOfUpd_ByRs"
'Dim mLnFld$: If Substract_Lst(mLnFld, ToStr_Flds(pRs.Fields), pLmPk) Then ss.A 1: GoTo E
'Dim mSet$: If RsSel(mSet, pRs, mLnFld) Then ss.A 1: GoTo E
'Dim mCndn$: If RsSel(mCndn, pRs, pLmPk$, , " and ") Then ss.A 1: GoTo E
'oSqlUpd = ToSql_Upd(TarTn, mSet, mCndn)
'Exit Function
'R: ss.R
'E:
End Function

Function SqlStrOfUpd_ByRsUlSrc(oSqlUpd$, pRsUlSrc As DAO.Recordset, pNmtHost$) As Boolean
'Aim: Build {oSql} to get data from Host by referring {pRsUlSrc}
'{pRsUlSrc} fmt: First N field is PK, Then [Changed], Then pair [xxx], [New xxx]
Const cSub$ = "SqlStrOfUpd_ByRsUlSrc"
Dim mSet$, mCndn$, mA$
Dim J%: For J = 0 To Fct.MinInt(10, pRsUlSrc.Fields.Count - 1)
    If pRsUlSrc.Fields(J).Name = "Changed" Then
        Dim I%: For I = J + 2 To pRsUlSrc.Fields.Count - 1 - 5 Step 2 'Skip 5 columns at end
            Dim mFld As DAO.Field: Set mFld = pRsUlSrc.Fields(I)
            If Not IsNull(mFld.Value) Then
                If Left(mFld.Name, 4) <> "New " Then ss.A 1, "The I-th field is not beging [New ]", , "I,NmFld", I, mFld.Name: GoTo E
                If Join_NmV(mA, Mid(mFld.Name, 5), mFld.Value) Then ss.A 2, "The I-th field cannot build 'Set xx=xx'", , "I,NmFld", I, mFld.Name: GoTo E
                mSet = Add_Str(mSet, mA)
            End If
        Next
        Exit For
    End If
    With pRsUlSrc.Fields(J)
        If Join_NmV(mA, .Name, .Value) Then ss.A 1: GoTo E
    End With
    mCndn = Add_Str(mCndn, mA, " and ")
Nxt:
Next
If mSet = "" Then ss.A 3, "mSet should be blank": GoTo E
If mCndn = "" Then ss.A 4, "mCndn should be blank": GoTo E
oSqlUpd = Fmt_Str("Update {0} set {1} where {2}", pNmtHost, mSet, mCndn)
Exit Function
R: ss.R
E:
End Function

Function SqlStrOfUpd_ByRsUlSrc__Tst()
Const cSub$ = "SqlStrOfUpd_ByRsUlSrc_Tst"
Dim mRs As DAO.Recordset, mSql$
Dim mRslt As Boolean, mCase As Byte: mCase = 1
Select Case mCase
Case 1
    If False Then
        If TblCrt_ForEdtTbl("tblUsr", 1) Then ss.A 1: GoTo E
    End If
    Set mRs = CurrentDb.OpenRecordset("Select * from tmpEdt_tblUsr")
    With mRs
        While Not .EOF
            mRslt = SqlStrOfUpd_ByRsUlSrc(mSql, mRs, "tmpEdt_tblUsr") ', mAmFm, mAmTo)
            .MoveNext
        Wend
        .Close
    End With
Case 2
End Select
Exit Function
R: ss.R
E:
End Function

Function SqlStrOfUpd1$(Tar$, Src$, NKeyFld%, Optional NKeyFldRmv% = 0, Optional A As database)
'Aim: Add/Upd {Tar} by {Src}.  Both has same {NKeyFld} of PK.  All fields in {Src} should all be found in {Tar}
'     If NKeyFldRmv>0 then some record in {Tar} will be remove if they does not exist in Src having first {NKeyFldRmv} as the matching keys
'     Example, Tar & Src: a,b,c, x,y,z
'              NKeyFld   : 3
'              NKeyFld   : 2
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
Dim OJn$, OSet$
    Dim KeyAy$(), SetAy$()
    Dim Fny$(): Fny = TblFny(Src, , A)
    Dim T2(): T2 = AyCut(Fny, CInt(NKeyFld))
    KeyAy = T2(0)
    SetAy = T2(1)
    OJn = StrExpand("t.{?}=s.{?}", KeyAy, , " and ")
    OSet = StrExpand("t.{?}=s.{?}", SetAy)

SqlStrOfUpd1 = FmtQQ("Update [?] t inner join [?] s on ? set ?", Tar, Src, OJn, OSet)
End Function

Function SqlStrOrdBy$(pOrdBy$)
If pOrdBy = "" Then Exit Function
SqlStrOrdBy = " Order by " & pOrdBy
End Function

Function SqlStrQuoteVar$(V)
Dim O$
Select Case VarType(V)
Case VbVarType.vbString: O = FmtQQ("'?'", V)
Case VbVarType.vbDate: O = FmtQQ("#?#", Format(V, "YYYY-DD-MM hh:mm:ss AM/PM"))
Case VbVarType.vbNull: O = "null"
Case Else: O = V
End Select
SqlStrQuoteVar = O
End Function

Function SqlStrQuoteVar1$(V)
Dim Ty As VbVarType: Ty = VarType(V)
If Ty = vbEmpty Then Exit Function
If Ty = vbNull Then SqlStrQuoteVar1 = "Null": Exit Function
If (Ty And vbArray) <> 0 Then Er "Given V is an array, cannot SqlStrQuoteVal"
Select Case VarSimTy(V)
Case eTypSim_Bool, eTypSim_Num: SqlStrQuoteVar1 = V
Case eTypSim_Dte: SqlStrQuoteVar1 = "#" & V & "#"
Case eTypSim_Str: SqlStrQuoteVar1 = Quote(Replace(V, "'", "''"), "'")
Case Else: Er "Unexpected {Ty}-of-V", TypeName(V)
End Select
End Function

Function SqlStrQuoteVar1__Tst()
Dim mV
mV = Now: Debug.Print SqlStrQuoteVar1(mV)
mV = 11: Debug.Print SqlStrQuoteVar1(mV)
mV = "12""34": Debug.Print SqlStrQuoteVar1(mV)
mV = "12'34": Debug.Print SqlStrQuoteVar1(mV)
mV = 12323&: Debug.Print SqlStrQuoteVar1(mV)
mV = 12323@: Debug.Print SqlStrQuoteVar1(mV)
mV = 12323!: Debug.Print SqlStrQuoteVar1(mV)
mV = 12323#: Debug.Print SqlStrQuoteVar1(mV)
mV = CByte(1): Debug.Print SqlStrQuoteVar1(mV)
mV = Null: Debug.Print SqlStrQuoteVar1(mV)
End Function

Function SqlStrToAnt(oAnt$(), Sql$) As Boolean
Const cSub$ = "SqlStrToAnt"
Dim mAnKw$(), mAyTypKw() As Byte, mAyKwLen() As Byte
If SqlStrToKwAy(mAnKw, mAyTypKw, mAyKwLen, Sql) Then Stop: GoTo E
Dim J%
Clr_Ays oAnt
For J = 0 To Siz_Ay(mAyTypKw) - 1
    If mAyTypKw(J) = 1 Then
        Dim mAnt$(): If Brk_Lnt(mAnt, Mid(mAnKw(J), mAyKwLen(J))) Then Stop: GoTo E
        Add_AyAtEnd oAnt, mAnt
    End If
Next
Exit Function
E: SqlStrToAnt = True
End Function

Function SqlStrToAnt__Tst()
Const cFt$ = "c:\aa.csv"
If Dlt_Fil(cFt) Then Stop: GoTo E
Dim mF As Byte: mF = FreeFile: Open cFt For Output As #mF
Print #mF, "Nmq,Lnt,Sql"
Dim mAnq$(), mAnt$(): If Fnd_Anq_ByLik(mAnq, "q*") Then Stop: GoTo E
Dim J%
For J = 0 To Siz_Ay(mAnq) - 1
    Debug.Print mAnq(J)
    Dim mSql$: mSql = CurrentDb.QueryDefs(mAnq(J)).Sql
    If SqlStrToAnt(mAnt, mSql) Then Stop: GoTo E
    Dim I%
    For I = 0 To Siz_Ay(mAnt) - 1
        Write #mF, mAnq(J), ToStr_Ays(mAnt), mSql
    Next
Next
Close #mF
Dim mWb As Workbook: If Opn_Wb(mWb, cFt, , , True) Then Stop
Exit Function
E:
End Function

Function SqlStrToKwAy(oAnKw$(), oAyTypKw() As Byte, oAyKwLen() As Byte, Sql$) As Boolean
'Aim: break {Sql} to {oAnKw} with {oAyTypKw} & {oAyKwLen}.  TypKw:1=From;2=Into;9=Other
Const cSub$ = "SqlStrToKwAy"
Clr_Ays oAnKw
Clr_AyByt oAyTypKw
Clr_AyByt oAyKwLen
Dim mSql$: mSql = RTrim(Replace(Replace(Sql, vbLf, " "), vbCr, " "))
Dim P%, mTypKw As Byte, mKwLen As Byte: P = InStr_SqlKw%(mTypKw, mKwLen, mSql)
While P > 0
    Dim mKw$: mKw = Right(mSql, Len(mSql) - P + 1)
    Add_AyEle oAnKw, mKw
    Add_AyByt oAyKwLen, mKwLen
    Add_AyByt oAyTypKw, mTypKw
    mSql = Left(mSql, P - 1)
    P = InStr_SqlKw%(mTypKw, mKwLen, mSql)
Wend
End Function

Sub SqlStrToKwAy__Tst()
Const cFt$ = "c:\aa.csv"
If Dlt_Fil(cFt) Then Stop: GoTo E

Dim mF As Byte: mF = FreeFile: Open cFt For Output As #mF
Print #mF, "Nmq,TypKw,Kw,Lnt,CleanLnt,Sql"

Dim mAnq$(): If Fnd_Anq_ByLik(mAnq, "qry*") Then Stop: GoTo E

Dim J%
For J = 0 To Siz_Ay(mAnq) - 1
    Debug.Print mAnq(J)
    Dim mSql$, mAnKw$(), mAyTypKw() As Byte, mAyKwLen() As Byte
    mSql = CurrentDb.QueryDefs(mAnq(J)).Sql
    If SqlStrToKwAy(mAnKw, mAyTypKw, mAyKwLen, mSql) Then Stop: GoTo E
    Dim I%
    For I = 0 To Siz_Ay(mAnKw) - 1
        Dim mLnt$: mLnt = Mid(mAnKw(I), mAyKwLen(I))
        Dim mAnt$(): If Brk_Lnt(mAnt, mLnt) Then Stop: GoTo E
        Write #mF, mAnq(J), mAyTypKw(I), mAnKw(I), mLnt, ToStr_Ays(mAnt), mSql
    Next
Next
Close #mF
Dim mWb As Workbook: If Opn_Wb(mWb, cFt, , , True) Then Stop
Exit Sub
E:
End Sub

Function SqlStrWhere$(Where$)
If Where = "" Then Exit Function
SqlStrWhere = " Where " & Where
End Function

