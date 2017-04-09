Attribute VB_Name = "nDao_Db"
Option Compare Database
Option Explicit

Function DbAppa(A As database) As Access.Application
If IsNothing(A) Then Set DbAppa = Access.Application: Exit Function
If ObjPtr(A) = ObjPtr(Access.Application.DBEngine.Workspaces(0).Databases(0)) Then
    Set DbAppa = Access.Application
End If
End Function

Sub DbCls(A As database)
On Error Resume Next
A.Close
Set A = Nothing
End Sub

Function DbHasQry(QryNm$, Optional A As database) As Boolean
On Error GoTo R
Dim Nm$: Nm = DbNz(A).QueryDefs(QryNm).Name
DbHasQry = True
Exit Function
R:
End Function

Function DbHasRel(RelNm$, Optional A As database) As Boolean
On Error GoTo R
Dim Nm$: Nm = DbNz(A).Relations(RelNm).Name
DbHasRel = True
Exit Function
R:
End Function

Function DbHasTbl(T, Optional A As database) As Boolean
Dim S$: S = FmtQQ("Select Count(*) from MSysObjects where Name='?' and Type in (1,6)", T)
DbHasTbl = SqlInt(S, A) = 1
End Function

Function DbNew(Fb$, Optional Locale$ = dbLangGeneral) As database
Set DbNew = Application.DBEngine.CreateDatabase(Fb, Locale)
End Function

Function DbNz(A As database) As database
If IsNothing(A) Then
    Set DbNz = DAO.DBEngine.Workspaces(0).Databases(0)
Else
    Set DbNz = A
End If
End Function

Function DbNz__Tst()
Debug.Print Application.DBEngine.Workspaces(0).Databases.Count
Dim mDb As database
Set mDb = DbNz(mDb)
Debug.Print Application.DBEngine.Workspaces(0).Databases.Count
Stop
End Function

Function DbQny(Optional A As database) As String()
Dim O$(), Q As QueryDef, Nm$
For Each Q In DbNz(A).QueryDefs
    Nm = Q.Name
    If Not IsPfxAp(Nm, "MSys", "~") Then Push O, Nm
Next
DbQny = O
End Function

Function DbRelCrt(pNmRel$, TFm$, TTo$, pLmFld$ _
    , Optional pIsIntegral As Boolean = False, Optional pIsCascadeUpd As Boolean = False, Optional pIsCascadeDlt As Boolean = False, Optional A As database) As Boolean
'Aim: Create a relation. {pLmFld} is format of xx=yy,cc,dd=ee
Const cSub$ = "DbRelCrt"
On Error GoTo R
Dim mDb As database: Set mDb = DbNz(A)
If IsRel(pNmRel) Then ss.A 1: GoTo E
Dim mAm() As tMap: mAm = Get_Am_ByLm(pLmFld)
If Siz_Am(mAm) = 0 Then ss.A 3, "pLmFld given 0 siz Am()": GoTo E
On Error GoTo R
Dim mRelAtr As DAO.RelationAttributeEnum
If Not pIsIntegral Then mRelAtr = dbRelationDontEnforce
If pIsCascadeUpd Then mRelAtr = mRelAtr Or dbRelationUpdateCascade
If pIsCascadeDlt Then mRelAtr = mRelAtr Or dbRelationDeleteCascade
Dim mRel As DAO.Relation: Set mRel = mDb.CreateRelation(pNmRel, TFm, TTo, mRelAtr)
Dim J%
For J = 0 To Siz_Am(mAm) - 1
    With mAm(J)
        mRel.Fields.Append mRel.CreateField(.F1)
        mRel.Fields(.F1).ForeignName = .F2
    End With
Next
mDb.Relations.Append mRel
Exit Function
R: ss.R
E: DbRelCrt = True: ss.B cSub, cMod, "pNmRel, TFm, TTo, pLmFld, pIsIntegral, pIsCascadeUpd, pIsCascadeDlt", pNmRel, TFm, TTo, pLmFld, pIsIntegral, pIsCascadeUpd, pIsCascadeDlt
End Function

Function DbRelCrt__Tst()
DbRelCrt "xxx#xx", "0Rec", "1Rec", "x", True, True, True
End Function

Function DbRelCrt_FmTbl(T$) As Boolean
'Aim: Create Relation for each record in {T}: Fb,NmTbl,NmTblTo,RelNo,IsCascadeDlt,IsCascadeUpd,LmFld
Const cSub$ = "DbRelCrt_FmTbl"
If Chk_Struct_Tbl(T, "Fb,NmTbl,NmTblTo,RelNo,IsCascadeDlt,IsCascadeUpd,LmFld") Then ss.A 1: GoTo E
On Error GoTo R
Dim mNmt$: mNmt = Q_SqBkt(T)
Dim mAyFb$(): mAyFb = SqlSy("Select Distinct Fb from " & mNmt)
Dim J%
For J = 0 To Siz_Ay(mAyFb) - 1
    Dim mDb As database: If Opn_Db_RW(mDb, mAyFb(J)) Then ss.A 3: GoTo E
    Dim mRs As DAO.Recordset, mSql$
    mSql = Bld_SqlSel( _
        "NmTbl,NmTblTo,RelNo,IsCascadeDlt,IsCascadeUpd,LmFld" _
        , mNmt _
        , "Fb='" & mAyFb(J) & "'" _
        , "NmTbl,RelNo")
    If Opn_Rs(mRs, mSql) Then ss.A 4: GoTo E
    With mRs
        While Not .EOF
            If DbRelCrt(!NmTbl & "R" & Format(!RelNo, "00"), "$" & !NmTbl, "$" & !NmTblTo, !LmFld, True, !IsCascadeUpd, !IsCascadeDlt, mDb) Then ss.A 5: GoTo E
            .MoveNext
        Wend
        .Close
    End With
    Cls_Db mDb
Next
GoTo X
R: ss.R
E: DbRelCrt_FmTbl = True: ss.B cSub, cMod, "T", T
X:
    RsCls mRs
    Cls_Db mDb
End Function

Sub DbRunSql(Sql, Optional A As database)
DbNz(A).Execute Sql
End Sub

Sub DbRunSqlNmAv(SqlNm, Av(), Optional A As database)
Dim S$: S = FmtNmAv(SqlNm, Av)
DbNz(A).Execute S
End Sub

Function DbStru$(Optional A As database)
Dim Ly$(), D As database: Set D = DbNz(A)
Dim T$(): T = DbTny(D)
Ly = AyMapInto(T, Ly, "TblStru", D)
DbStru = LyJn(Ly)
End Function

Function DbTblDefAy(D As database) As TableDef()
Dim O() As TableDef
Dim I As TableDef, J%
For Each I In D.TableDefs
    PushObj O, I
Next
DbTblDefAy = O
End Function

Function DbTblFldDt(A As database) As Dt
Dim D As database: Set D = DbNz(A)
Dim T$(): T = DbTny(D)
Dim DrAy()
DbTblFldDt = DtNew(LvsSplit("Tbl Fld"), DrAy)
End Function

Function DbTmp(Optional Locale$ = dbLangGeneral, Optional Pfx$, Optional SubFdr$) As database
Set DbTmp = DbNew(TmpFb, Locale)
End Function

Function DbTny(Optional A As database) As String()
Dim O$(), T As TableDef, Nm$
For Each T In DbNz(A).TableDefs
    Nm = T.Name
    If Not IsPfxAp(Nm, "MSys", "~") Then Push O, Nm
Next
DbTny = O
End Function

Function DbTxt(Pth$) As database
'Aim: Open {pDir} as a database by referring all *.txt as table
'Note: Schema.ini in {pDir} will be used if exist.  See Fdf2Schema() about Schema.ini
Dim Fb$
Stop
Set DbTxt = G.gDbEng.OpenDatabase(Fb, False, False, "Text;Database=" & Pth)
End Function

