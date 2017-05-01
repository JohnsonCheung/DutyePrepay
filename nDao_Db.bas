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

Function DbOfTxtPth(TxtPth$) As database
'Aim: Open {pDir} as a database by referring all *.txt as table
'Note: Schema.ini in {pDir} will be used if exist.  See Fdf2Schema() about Schema.ini
Dim Fb$
Stop
Set DbOfTxtPth = DBEngine.OpenDatabase(Fb, False, False, "Text;Database=" & TxtPth)
End Function

Function DbPth$(Optional A As database)
DbPth = FfnPth(DbNz(A).Name)
End Function

Function DbQny(Optional A As database) As String()
Dim O$(), Q As QueryDef, Nm$
For Each Q In DbNz(A).QueryDefs
    Nm = Q.Name
    If Not IsPfxAp(Nm, "MSys", "~") Then Push O, Nm
Next
DbQny = O
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

Sub DbWrtFx_wTp(Tvnstr, Fx$, FxTp$, _
    Optional WsNmPfx$, _
    Optional WsNmSfx$, _
    Optional NoExpTim As Boolean, _
    Optional A As database)
'Aim: Export all tables/queries in {TqnStr} to {Fx} with {pWsNmPfx/pWsNmSfx} added to each Ws (ie ws name will be pPfx + Nmtq + pSfx}.
'     "Note to Nmtq": if Nmtq is in format of xxx_Oup_yyy or #@yyy, yyy will be use
'FfnAsstExist FxTp
'FfnAsstNotExist Fx
'FfnCpy FxTp, Fx
'Dim Db As database: Set Db = DbNz(A)
'Dim Ny$(): Ny = NmstrBrk(Tvnstr)
'Dim OWb As Workbook, mToBeDelete$: mToBeDelete = ""
'Set OWb = FxWb(Fx)
'Dim I
'For Each I In Ny
'    Dim Ws As Worksheet
'    Dim mWsNmTar$: mWsNmTar = WsNmPfx & Cut_Aft(Cut_Aft(Cut_Aft(mAntq(I), "_Oup_"), "#@"), "@") & pWsNmSfx
'    If Fnd_Ws(mWs, mWb, mWsNmTar, True) Then
'        If Add_Ws(mWs, mWb, mWsNmTar) Then Add_AyEle mAntqErr, mAntq(I): GoTo Nxt
'        If Exp_Nmtq2Ws(mAntq(I), mWs, SrcFb) Then Add_AyEle mAntqErr, mAntq(I): GoTo Nxt
'    Else
'        If Exp_Nmtq2Ws_wFmt_ByCpyRs(mAntq(I), mWs.Range("A5"), SrcFb, pNoExpTim) Then Add_AyEle mAntqErr, mAntq(I): GoTo Nxt
'    End If
'Nxt:
'Next
'If mToBeDelete$ <> "" Then Dlt_Ws_InWb mWb, mToBeDelete
'If Sz(mAntqErr) > 0 Then ss.A 3, "These tables cannot be exported: " & Join(mAntqErr, ","): GoTo E
'WbCls OWb
End Sub

