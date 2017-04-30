Attribute VB_Name = "nDao_Qry"
Option Compare Database
Option Explicit

Function Qny(Optional Lik$ = "*", Optional A As database) As String()
Qny = AyLik(ObjAyNy(QryAy(A)), Lik)
End Function

Function Qry(Qn, Optional A As database) As QueryDef
Set Qry = DbNz(A).QueryDefs(Qn)
End Function

Function QryAy(Optional A As database) As QueryDef()
Dim O() As QueryDef, Q As QueryDef
For Each Q In DbNz(A).QueryDefs
    If IsPfx(Q.Name, "~") Then GoTo Nxt
    PushObj O, Q
Nxt:
Next
QryAy = O
End Function

Sub QryCpy(FmPfx$, ToPfx$)
QryDrp_ByPfx ToPfx
Dim iQry As QueryDef
Dim mLstPart$
For Each iQry In CurrentDb.QueryDefs
    If Left(iQry.Name, Len(FmPfx)) = FmPfx Then
        mLstPart = Mid$(iQry.Name, Len(FmPfx) + 1)
        Debug.Print "Fm:" & iQry.Name & "  To:" & ToPfx & mLstPart
        Call DoCmd.CopyObject(, ToPfx & mLstPart, AcObjectType.acQuery, FmPfx & mLstPart)
    End If
Next
End Sub

Function QryCrt(QryNm$, Optional Sql$ = "", Optional A As database) As Boolean
Const cSub$ = "QryCrt"
On Error GoTo R
Dim mDb As database: Set mDb = DbNz(A)
Dim mNmq$: mNmq = Rmv_SqBkt(QryNm)
With mDb
    If IsQry(mNmq, mDb) Then
        If .QueryDefs(mNmq).Type = DAO.QueryDefTypeEnum.dbQSQLPassThrough Then
            .QueryDefs.Delete (mNmq)
            .CreateQueryDef mNmq
        End If
    Else
        .CreateQueryDef mNmq
    End If
    Dim mQry As DAO.QueryDef: Set mQry = .QueryDefs(mNmq)
    If Sql <> "" Then mQry.Sql = Sql
    .QueryDefs.Refresh
End With
Exit Function
R: ss.R
E:
End Function

Sub QryCrt_ByDSN(QryNm$, Sql$, Dsn$, IsRetRec As Boolean, Optional A As database)
Dim D As database: Set D = DbNz(A)
QryCrt QryNm, Sql, A
Dim Q As DAO.QueryDef: Set Q = D.QueryDefs(QryNm)
Q.Connect = "ODBC;DSN=" & Dsn & ";"
QrySetPrp_Bool Q, "ReturnsRecords", IsRetRec
QrySetPrp Q, "ODBCTimeout", SysCfg_OdbcTimeOut
End Sub

Function QryCrt_ByDSN__Tst()
'mSql = "Select SUM(Case When ICLAS IN ('57','07') Then 1 Else 0 end) AA , SUM(Case When ICLAS IN ('14','64') Then 1  Else 0 end) BB from IIC"
'If QryCrt_ByDSN("qry", mSql, "FEPROD_RBPCSF") Then Stop
Shw_DbgWin
Debug.Print "----ReturnsRecords True------"
QryCrt_ByDSN "xxxy", "Update YY SET ICDES='11' WHERE ICLAS='06'", "FETEST_QGPL", True
Debug.Print ToStr_Prps(CurrentDb.QueryDefs("XXX").Properties, vbCrLf)
Debug.Print "----ReturnsRecords False ------"
QryCrt_ByDSN "xxx", "Update YY SET ICDES='11' WHERE ICLAS='06'", "FETEST_QGPL", False
Debug.Print ToStr_Prps(CurrentDb.QueryDefs("XXX").Properties, vbCrLf)
Debug.Print "----ReturnsRecords True------"
QryCrt_ByDSN "xxx", "Update YY SET ICDES='11' WHERE ICLAS='06'", "FETEST_QGPL", True
Debug.Print ToStr_Prps(CurrentDb.QueryDefs("XXX").Properties, vbCrLf)
Debug.Print "----ReturnsRecords False ------"
QryCrt_ByDSN "xxx", "Update YY SET ICDES='11' WHERE ICLAS='06'", "FETEST_QGPL", False
Debug.Print ToStr_Prps(CurrentDb.QueryDefs("XXX").Properties, vbCrLf)
End Function

Function QryCrt_FmTbl(T) As Boolean
'Aim: Create all queries as defined in {T}: Fb,NmQry,Sql
Const cSub$ = "QryCrt_FmTbl"
If Chk_Struct_Tbl(CStr(T), "Fb,NmQry,Sql") Then ss.A 1: GoTo E
Dim mFbLas$, mRs As DAO.Recordset
If Opn_Rs(mRs, "Select * from [" & Rmv_SqBkt(CStr(T)) & "] order by Fb,NmQry") Then ss.A 2: GoTo E
With mRs
    While Not .EOF
        If mFbLas <> !Fb Then
            mFbLas = !Fb
            Dim mDb As database: Cls_Db mDb: If Opn_Db_RW(mDb, mFbLas) Then ss.A 2: GoTo E
        End If
        If QryCrt(!NmQry, !Sql, mDb) Then ss.A 3: GoTo E
        .MoveNext
    Wend
End With
GoTo X
R: ss.R
E:
X:
    Cls_Db mDb
    RsCls mRs
End Function
Function QryLy(QnStr$, Optional SqlSubStr$, Optional A As database) As String()
'Dim L%: L = Len(QryNmPfx)
'Dim iQry As QueryDef: For Each iQry In CurrentDb.QueryDefs
'    If Left(iQry.Name, L) = QryNmPfx Then If InStr(iQry.Sql, Sql_SubString) > 0 Then Debug.Print ToStr_TypQry(iQry.Type), iQry.Name
'Next
'End Function
'Function Lst_QryPrm_ByPfx(QryNmPfx$, Optional pFno As Byte = 0) As Boolean
'Dim L%: L = Len(QryNmPfx)
'Dim iQry As QueryDef: For Each iQry In CurrentDb.QueryDefs
'    If Left(iQry.Name, L) = QryNmPfx Then
'        If iQry.Parameters.Count > 0 Then
'            Prt_Str pFno, iQry.Name & "-----(Param)------>"
'            Dim iPrm As DAO.parameter
'            For Each iPrm In iQry.Parameters
'                Prt_Str pFno, iPrm.Name
'            Next
'            Prt_Ln pFno
'        End If
'    End If
'Next
End Function

Function QryCrt_FmTbl__Tst()
TblCrt_ByFldDclStr "#FBQry", "Fb Text 255,NmQry Text 50,Sql Memo"
If Run_Sql("Insert into [#FBQry] values ('C:\Tmp\aa.mdb','qry1','select * from Tbl1')") Then GoTo E
If Run_Sql("Insert into [#FBQry] values ('C:\Tmp\aa.mdb','qry2','select * from Tbl1')") Then GoTo E
If Run_Sql("Insert into [#FBQry] values ('C:\Tmp\aa.mdb','qry3','select * from Tbl1')") Then GoTo E
If Run_Sql("Insert into [#FBQry] values ('C:\Tmp\aa.mdb','qry4','select * from Tbl1')") Then GoTo E
FbNew "c:\Tmp\aa.mdb"
If QryCrt_FmTbl("#FBQry") Then GoTo E
G.gAcs.OpenCurrentDatabase "c:\tmp\aa.mdb"
G.gAcs.Visible = True
Exit Function
E:
End Function

Sub QryCrtSubDtaSheet(MstQn$, ChdQn$, MstFnStr$, Optional ChdFnStr$, Optional A As database)
Dim O As QueryDef: Set O = DbNz(A).QueryDefs(MstQn)
Dim OMst$
    OMst = AyJnComma(NmstrBrk(MstFnStr))
Dim OChd$
    OChd = IIf(ChdFnStr = "", MstFnStr, ChdFnStr)
    OChd = AyJnComma(NmstrBrk(OChd))
QrySetPrp O, "SubdatasheeQname", ChdQn
QrySetPrp O, "LinkChildFields", OChd
QrySetPrp O, "LinkMasterFields", OMst
End Sub

Sub QryDmp(Optional Lik$ = "*", Optional A As database)
Dim J%, D As database
Set D = DbNz(A)
For J = 0 To CurrentDb.QueryDefs.Count - 1
    If CurrentDb.QueryDefs(J).Name Like Lik Then
        Debug.Print D.QueryDefs(J).Name
        Debug.Print D.QueryDefs(J).Sql
        Debug.Print
    End If
Next
End Sub

Sub QryDrp(Qn, Optional A As DAO.database)
DbNz(A).QueryDefs.Delete Qn
End Sub

Sub QryDrp_ByLik(Lik$, Optional A As database)
Dim Q$(): Q = Qny(Lik, A)
QryDrp_ByQny Q, A
End Sub

Function QryDrp_ByPfx(Pfx, Optional A As database) As Boolean
QryDrp_ByQny Qny(Pfx & "*", A)
End Function

Sub QryDrp_ByQny(Qny$(), Optional A As database)
If AyIsEmpty(Qny) Then Exit Sub
Dim I
For Each I In Qny
    QryDrp I, A
Next
End Sub

Sub QryDrpPrp(Q As QueryDef, PrpNm$)
If QryHasPrp(Q, PrpNm) Then Q.Properties.Delete PrpNm
End Sub

Sub QryExpToFb(Qry_or_Tbl_Nm$, TarFb$, Optional TarTn$ = "", Optional SrcFb$ = "", Optional OvrWrt As Boolean = False)
'Aim: Export {Qry_or_Tbl_Nm} in {p.FbSrc} to table {p.NmtTar} in {p.FbTar}.  {Nmt2Mdb} will be created if not exist
Const cSub$ = "Exp_Nmq2Mdb"
On Error GoTo R
If VBA.Dir(TarFb) = "" Then FbNew TarFb
Dim mNmtTar$: mNmtTar = IIf(TarTn = "", Qry_or_Tbl_Nm, TarTn)
On Error GoTo R
Dim mIn_FbSrc$: If SrcFb <> "" Then mIn_FbSrc = " in '" & SrcFb & CtSngQ
Dim mSql$: mSql = Fmt_Str("select * into {0} in '{1}' from {2}{3}", mNmtTar, TarFb, Qry_or_Tbl_Nm, mIn_FbSrc)
If Run_Sql(mSql) Then ss.A 2: GoTo E
Exit Sub
R: ss.R
E:
End Sub

Function QryExpToFb__Tst()
Dim mCase As Byte, mNmtq$, mFbTar1$, mFbTar2$, mNmtTar1$, mNmtTar2$
mFbTar1 = "c:\aa.mdb"
mFbTar2 = "c:\bb.mdb"
If Dlt_Fil(mFbTar1) Then Stop
If Dlt_Fil(mFbTar2) Then Stop
For mCase = 1 To 2
    Select Case mCase
    Case 1: mNmtq = "qryAllBrand"
    Case 2: mNmtq = "query1"
    End Select
    QryExpToFb mNmtq, mFbTar1
Next
QryExpToFb "qryAllBrand", mFbTar2, mNmtTar1
QryExpToFb "query1", mFbTar2, mNmtTar2

G.gAcs.OpenCurrentDatabase mFbTar1
G.gAcs.Visible = True
Dim mAcs As New Access.Application
mAcs.OpenCurrentDatabase mFbTar2
mAcs.Visible = True
Stop
End Function

Function QryHasPrp(Q As QueryDef, PrpNm$) As Boolean
QryHasPrp = PrpIsExist(PrpNm, Q.Properties)
End Function

Sub QryOpn(QryNm)
DoCmd.OpenQuery QryNm, , acReadOnly
End Sub

Function QryRenPfx(FmPfx$, ToPfx$) As Boolean
Dim iQry As QueryDef
Dim L%: L = Len(FmPfx)
For Each iQry In CurrentDb.QueryDefs
    If Left(iQry.Name, L) = FmPfx Then
        Debug.Print "Replacing Qry ... "; iQry.Name
        iQry.Name = ToPfx & Mid$(iQry.Name, L + 1)
    End If
Next
End Function

Function QryRenPfxSet(pQryPfx$, pBegNum As Byte, pEndNum As Byte, pToNum As Byte) As Boolean
If pToNum = pBegNum Then MsgBox "pToNum must <> pBegNum": Exit Function
If pEndNum < pBegNum Then MsgBox "pEndNum must > pBegNum": Exit Function
Dim J%
If pToNum > pBegNum Then
    For J = pEndNum To pBegNum Step -1
        QryRenPfx pQryPfx & "_" & VBA.Format(J, "00"), pQryPfx & "_" & VBA.Format(J + pToNum - pBegNum, "00")
    Next
Else
    For J = pBegNum To pEndNum
        QryRenPfx pQryPfx & "_" & VBA.Format(J, "00"), pQryPfx & "_" & VBA.Format(J + pToNum - pBegNum, "00")
    Next
End If
End Function

Sub QrySetPrp(Q As QueryDef, PrpNm$, V)
If VarIsBlank(V) Then
    QryDrpPrp Q, PrpNm
    Exit Sub
End If

If QryHasPrp(Q, PrpNm) Then
    Q.Properties(PrpNm).Value = V
Else
    Q.Properties.Append Q.CreateProperty(PrpNm, VarDaoTy(V), V)
End If
End Sub

Function QrySetPrp_Bool(pQry As QueryDef, PrpNm$, V As Boolean) As Boolean
On Error GoTo R
pQry.Properties(PrpNm).Value = V
Exit Function
R: ss.R
    Dim mPrp As DAO.Property: Set mPrp = pQry.CreateProperty(PrpNm, DAO.DataTypeEnum.dbBoolean, V)
    pQry.Properties.Append mPrp
End Function

Sub QrySetRmk(QryNm, Rmk$, Optional A As database)
QrySetPrp Qry(QryNm, A), "Description", Rmk
End Sub

Sub QryShw(QryNm$)
DoCmd.OpenQuery QryNm, acViewDesign
End Sub

Function QrySql$(QryNm, Optional A As database)
QrySql$ = DbNz(A).QueryDefs(QryNm).Sql
End Function
