Attribute VB_Name = "nDao_nCrt_Tbl"
Option Compare Database
Option Explicit

Sub TblCrt(T, SqlFldLst$, Optional A As database)
DbRunSql SqsOfCrt(T, SqlFldLst), A
End Sub

Sub TblCrt_ByFldAy(T, FldAy() As DAO.Field _
    , Optional NPk As Byte = 0 _
    , Optional TblAtr As DAO.TableDefAttributeEnum _
    , Optional PkNoAutoInc As Boolean _
    , Optional A As database)
'Aim: Delete then Create {T} in {A} by {FldAy} with {TblAtr}.
Dim D As database: Set D = DbNz(A)
    
Dim OFldAy() As DAO.Field
    OFldAy = FldAy
    If NPk = 1 And Not PkNoAutoInc Then
        If OFldAy(0).Type = dbLong Then FldAy(0).Attributes = DAO.FieldAttributeEnum.dbAutoIncrField
    End If
    
Dim OTbl As DAO.TableDef: Set OTbl = D.CreateTableDef(T, TblAtr)
   
Dim OIdx As DAO.Index
    If NPk > 0 Then
        Set OIdx = OTbl.CreateIndex("PrimaryKey")
        OIdx.Unique = True
        OIdx.Primary = True
        Dim J%
        For J = 0 To NPk - 1
            OIdx.Fields.Append OFldAy(J)
        Next
    End If
'----
Dim F
For Each F In OFldAy
    OTbl.Fields.Append F
Next

If NPk > 0 Then OTbl.Indexes.Append OIdx

TblDrp T, D
With D.TableDefs
    .Append OTbl
    .Refresh
End With
End Sub

Sub TblCrt_ByFldDclStr(T, FldDclStr$ _
        , Optional NPk As Byte = 0 _
        , Optional TblAtr As DAO.TableDefAttributeEnum = 0 _
        , Optional A As database)
'Aim: Delete then Create {A}!{T} by {FldDclStr} with {TblAtr}.
'     Format of FldDclStr is xxx Text 10,....
'     Note: xxx may be in xx^xx format.  ^ means for space
'       TEXT,CURRENCY,LONG,INT,BYTE,DATE,SINGLE,DOUBLE,MEMO,YESNO
Dim FldDclSy$(): FldDclSy = Split(FldDclStr, CtComma)
TblCrt_ByFldDclSy T, FldDclSy, NPk, TblAtr, , A
End Sub

Function TblCrt_ByFldDclStr__Tst()
'If Run_Sql("Create table aXa (bb NUMERIC)") Then Stop
Dim mFmLoFld$, mNmt$
Dim Db As database: If Crt_Db(Db, "c:\tmp\aa.Db", True) Then Stop
Dim mCase As Byte: mCase = 2
Select Case mCase
Case 1
    mNmt$ = "XX"
    mFmLoFld = "aa Long, bb Int, cc currency 4,TT TEXT 10"
Case 2
    mNmt = "MSysIMEXSpecs"
    mFmLoFld = "SpecName Text 64" & _
        ", SpecId Auto" & _
        ", DateDelim Text 2" & _
        ", DateFourDigitYear YesNo" & _
        ", DateLeadingZeros YesNo" & _
        ", DecimalPoint Text 2" & _
        ", DateOrder Int" & _
        ", FieldSeparator Text 2" & _
        ", FileType Int" & _
        ", SpecType Byte" & _
        ", StartRow Long" & _
        ", TextDelim Text 2" & _
        ", TimeDelim Text 2"
    TblCrt_ByFldDclStr mNmt, mFmLoFld, 1, 2, Db
    
End Select
Cls_Db Db
If Opn_CurDb(G.gAcs, "c:\tmp\aa.Db") Then Stop
G.gAcs.Visible = True
Stop
GoTo X
E:
X: Cls_CurDb G.gAcs
End Function

Sub TblCrt_ByFldDclSy(T, FldDclSy$() _
    , Optional NPk As Byte = 0 _
    , Optional TblAtr As DAO.TableDefAttributeEnum _
    , Optional PkNoAutoInc As Boolean _
    , Optional A As database)
Dim FldAy() As Field
    FldAy = AyMapInto(FldDclSy, FldAy, "FldDclStrFld")

TblCrt_ByFldAy T, FldAy, NPk, TblAtr, PkNoAutoInc, A
End Sub

Sub TblCrt_FmDSN_Nmt(T$, Dsn$, Optional SrcTn$, Optional A As database)
Dim Src$
    Src = IIf(SrcTn = "", T, SrcTn)
TblCrt_FmDSN_Sql T, Dsn, SqsOfSel(Src)
End Sub

Function TblCrt_FmDSN_Nmt__Tst()
TblCrt_FmDSN_Nmt "#IIC", "FEPROD_RBPCSF", "iic"
End Function

Sub TblCrt_FmDSN_Sql(T, Dsn$, Sql$, Optional A As database)
'Aim: Download Data to {T} in {TarFb} by {Sql} through {Dsn$}
Dim TmpQ$: TmpQ = TmpNm("Qry")
QryCrt_ByDSN TmpQ, Sql, Dsn, IsRetRec:=True
Dim S$: S = FmtQQ("Select * into [?] from [?]", T, TmpQ)
DbRunSql S, A
QryDrp TmpQ
End Sub

Function TblCrt_FmDSN_Sql__Tst()
Const cSub$ = "TblCrt_FmDSN_Sql_Tst"
Dim mDsn$, mSql$, mNmtTar$, mFbTar$
Dim mRslt As Boolean, mCase As Byte
Dim mNRec&, mDteBeg As Date, mDteEnd As Date
For mCase = 1 To 4
    Select Case mCase
    Case 1: mDsn = "FEPROD_RBPCSF": mSql = "Select * from IIC": mNmtTar = "IIC_Xls": mFbTar = "C:\aa.Db"
    Case 2: mDsn = "FEPROD_RBPCSF": mSql = "Select * from IIC": mNmtTar = "IIC_Txt": mFbTar = "C:\aa.Db"
    Case 3: mDsn = "FEPROD_RBPCSF": mSql = "Select * from IIC": mNmtTar = "IIC_Xls": mFbTar = ""
    Case 4: mDsn = "FEPROD_RBPCSF": mSql = "Select * from IIC": mNmtTar = "IIC_Txt": mFbTar = ""
    End Select
    TblCrt_FmDSN_Sql mNmtTar, mDsn, mSql
    Debug.Print mCase; "-----------------------"
    Debug.Print ToStr_LpAp(vbLf, "mRslt,mDsn,mSql,mNmtTar,mFbTar,mDteBeg,mDteEnd,mNRec", mRslt, mDsn, mSql, mNmtTar, mFbTar, mDteBeg, mDteEnd, mNRec)
Next
End Function

Sub TblCrt_FmDTF_Nmt(T, IP$, Lib$ _
    , Optional SrcT$ _
    , Optional IsByXls As Boolean _
    , Optional IsKeepDownloadFfn As Boolean _
    , Optional ONrec& _
    , Optional A As database)
'Aim: Create {TarTn} in {TarFb} from {pIP},{pLib},{T} by meaning DTF download through {pIsByXls} or by Text
Dim S$: S = SqsOfSel(IIf(SrcT = "", T, SrcT))
TblCrt_FmDTF_Sql T, IP, S, Lib, IsByXls, IsKeepDownloadFfn, ONrec, A
End Sub

Function TblCrt_FmDTF_Nmt__Tst()
Dim mNRec&, T$, IsByXls As Boolean

Dim Fb$
    Fb = TmpFb
    FbNew Fb
    
Dim Db As database
    Set Db = FbDb(Fb)
Dim J%
For J = 1 To 4
    Select Case J
        Case 1: T = "IIC_ByXls": IsByXls = True:
        Case 2: T = "IIC_ByTxt": IsByXls = False
        Case 3: T = "IIC_ByXls": IsByXls = True
        Case 4: T = "IIC_ByTxt": IsByXls = False
    End Select
    TblCrt_FmDTF_Nmt T, "192.168.103.14", "RBPCSF", "IIC", IsByXls, , mNRec, Db
    Debug.Print ToStr_LpAp(vbTab, "IsByXls, mNRec", IsByXls, mNRec)
Next
End Function

Sub TblCrt_FmDTF_Sql(T, IP$, Lib$, Sql$ _
    , Optional IsByXls As Boolean _
    , Optional IsKeepDownloadFfn As Boolean _
    , Optional ONrec& _
    , Optional A As database)
'Aim: Create {TarTn} in {TarFb} from {pIP},{pLib},{Sql} with time stamped & Rec count {oDteBeg,oDteEnd,oNRec&}.

Dim Dtf$: Dtf = TmpFil(".dtf", , Lib)
DtfCrt Dtf, Sql, IP, Lib, IsByXls, IsRun:=True, ONrec:=ONrec

Dim Ext$: Ext = IIf(IsByXls, ".xls", ".txt")
Dim F$: F = FfnRplExt(Dtf, Ext)
If IsByXls Then
    TblCrt_FmFx_n_FDF T, F, IsKeepDownloadFfn, A
Else
    TblCrt_FmFt_n_FDF T, F, IsKeepDownloadFfn, A
End If
End Sub

Function TblCrt_FmDTF_Sql__Tst()
Const cSub$ = "TblCrt_FmDTF_Sql_Tst"
Dim mNRec&, mNmt$, mDteBeg As Date, mDteEnd As Date, mIsByXls As Boolean, mRslt
Dim mFbTar$
Dim mCase As Byte
For mCase = 3 To 3
    Select Case mCase
        Case 1: mNmt = "IIC_ByXls": mIsByXls = True: mFbTar = "c:\aa.Db"
        Case 2: mNmt = "IIC_ByTxt": mIsByXls = False: mFbTar = "c:\aa.Db"
        Case 3: mNmt = "IIC_ByXls": mIsByXls = True: mFbTar = ""
        Case 4: mNmt = "IIC_ByTxt": mIsByXls = False: mFbTar = ""
    End Select
    TblCrt_FmDTF_Sql "192.168.103.13", "Select * from IIC where ICLAS='07'", mNmt, mFbTar, "BPCSF", mIsByXls, mNRec
    Debug.Print ToStr_LpAp(vbLf, "mRslt, mFbTar, mIsByXls, mDteBeg, mDteEnd, mNRec", mRslt, mFbTar, mIsByXls, mDteBeg, mDteEnd, mNRec)
Next
End Function

Sub TblCrt_FmFt_n_FDF(T, Ft$, Optional KeepFt As Boolean, Optional A As database)
'Aim: Create a table {T} in Db-{A} by import a text file {Ft} by buiding a schema.ini Fm {Fdf}-file
Dim D As database: Set D = DbNz(A)

Dim Fdf$
    Fdf = FfnRplExt(Ft, ".Fdf")

'#2 Build Schema.ini in {pDir}
FdfWrtSchemaIni Fdf

Dim Pth$
    Pth = FfnPth(D.Name)
    
Dim SrcT$
    SrcT = FfnFnn(Ft) & "#Txt"

Dim Sel$:
    Sel = AyJnComma(FdfFny(Fdf))

Dim CnnStr$
    CnnStr = "Text;Database=" & Pth

Dim S$
    S = FmtQQ("Select ? into [?] from [?] in '' [?]", Sel, T, SrcT, CnnStr)

DbRunSql S, D

'#4 Dlt Txt, Fdf & Schema.ini if success
If KeepFt Then
    FfnDlt Ft
    FfnDlt Fdf
End If
FfnDlt FfnPth(Ft) & "Schema.ini"
End Sub

Function TblCrt_FmFt_n_FDF__Tst()
Const Dtf$ = "C:\Tmp\IIC.dtf"
Const T$ = ">IIC"
DtfCrt Dtf, "Select * from IIC", "192.168.103.14", , , True
TblCrt_FmFt_n_FDF T, FfnRplExt(Dtf, ".txt")
End Function

Sub TblCrt_FmFx(T, Fx$, Optional IsKeepFx As Boolean, Optional A As database)
'Aim: Create a table {TarTn} in {TarFb} by import an Xls file {pFx} with referring CutExt{pFx}.Fdf
Dim D As database: Set D = DbNz(A)

FfnAsstExist Fx, "TblCrt_FmFx_n_Fdf"

'Import
Dim CnnStr$: CnnStr = CnnStr_Xls(Fx)
Dim SrcT$
Dim S$: S = FmtQQ("Select * into [?] from [?] in '' [?]", T, SrcT, CnnStr)
DbRunSql S, D
If Not IsKeepFx Then
    FfnDltIfExist Fx
End If
End Sub

Sub TblCrt_FmFx_n_FDF(T, Fx$, Optional KeepFx As Boolean, Optional A As database)
'Aim: Create a table {T} in Db-{A} by import a text file {Fx} by buiding a schema.ini Fm {Fdf}-file
Dim D As database: Set D = DbNz(A)

Dim Fdf$
    Fdf = FfnRplExt(Fx, ".Fdf")

Dim Pth$
    Pth = FfnPth(D.Name)
    
Dim SrcT$
    SrcT = FfnFnn(Fx)

Dim Sel$:
    Sel = AyJnComma(FdfFny(Fdf))

Dim CnnStr$
    CnnStr = CnnStrFx(Fx)
Dim S$
    S = FmtQQ("Select ? into [?] from [?] in '' [?]", Sel, T, SrcT, CnnStr)

DbRunSql S, D

'#4 Dlt Txt, Fdf & Schema.ini if success
If KeepFx Then
    FfnDlt Fx
    FfnDlt Fdf
End If
FfnDlt FfnPth(Fx) & "Schema.ini"
End Sub

Function TblCrt_FmFx_n_FDF__Tst()
Const Dtf$ = "C:\Tmp\IIC.dtf"
Const Fx$ = "C:\Temp\IIC.xls"
Const T$ = "IIC"
DtfCrt Dtf, "Select * from IIC where ICLAS='xx'", "192.168.103.14", , , True, True
TblCrt_FmFx_n_FDF T, Fx
End Function

Function TblCrt_FmLnk(T, TSrc$, pCnn$, Optional A As database) As Boolean
'Aim: Create {T} in {pInDb} by linking {TSrc} using {pCnn}
Dim D As database: Set D = DbNz(A)
TblDrp T, D
Dim mTbl As New DAO.TableDef
With mTbl
    .Connect = pCnn
    .Name = T
    .SourceTableName = TSrc
    D.TableDefs.Append mTbl
End With
End Function

Function TblCrt_FmLnk__Tst()
Dim mNmt$:      mNmt = "A1"
Dim mNmtSrc$:   mNmtSrc = "a1.txt"
Dim mCnn$:      mCnn = "Text;DSN=A1;FMT=Fixed;HDR=NO;IMEX=2;CharacterSet=20127;DATABASE=c:\;TABLE=a1#txt"
Dim Db As database: If Crt_Db(Db, "c:\aa.Db", True) Then Stop
If TblCrt_FmLnk(mNmt, mNmtSrc, mCnn, Db) Then Stop
End Function

Function TblCrt_FmLnkAs400Dsn(T, Optional pLib$ = "RBPCSF", Optional pAs400Dsn$ = "FEPROD_RBPCSF", Optional TNew$ = "", Optional pInDb As database) As Boolean
'Aim: Create NonBlank({TNew},{pLib}_{T}) in {pInDb} by linking {T} through {pAs400Dsn}.  Dsn must use *SQL Naming Convertion, ie
Const cSub$ = "TblCrt_FmLnkAs400Dsn"
Dim mNmt$: mNmt = NonBlank(TNew, pLib & "_" & T)
Dim mCnn$: mCnn = Fmt_Str("ODBC;DSN={0};", pAs400Dsn)
Dim mNmtSrc$: mNmtSrc = pLib & "." & T
TblCrt_FmLnkAs400Dsn = TblCrt_FmLnk(mNmt, mNmtSrc, mCnn, pInDb)
Exit Function
R: ss.R
E:     Debug.Print "<--- Cannot link"
End Function

Function TblCrt_FmLnkAs400Dsn__Tst()
If TblCrt_FmLnkAs400Dsn("IIC", , , "xx") Then Stop
End Function

Function TblCrt_FmLnkCsv(pFfnCsv$, Optional TNew$ = "", Optional A As database) As Boolean
Const cSub$ = "TblCrt_FmLnkCsv"
Dim Db As database: Set Db = DbNz(A)
Dim mNmtNew$: If TNew = "" Then mNmtNew = Fct.Nam_FilNam(pFfnCsv) Else mNmtNew = TNew
Dlt_Tbl mNmtNew, Db
Dim mTbl As New DAO.TableDef
On Error GoTo R
With mTbl
    Dim mDir$, mFnn$, mExt$
    Call Brk_Ffn_To3Seg(mDir, mFnn, mExt, pFfnCsv)
    .Connect = Fmt_Str("Text;DSN=Import Link Specification;FMT=Delimited;HDR=NO;IMEX=2;CharacterSet=936;DATABASE={0};TABLE={1}#{2}", mDir, mFnn, Mid(mExt, 2))
    .Name = mNmtNew
    .SourceTableName = mFnn & mExt
    Db.TableDefs.Append mTbl
End With
On Error GoTo 0
Exit Function
R: ss.R
E:
'Text;DSN=Import Link Specification;FMT=Delimited;HDR=NO;IMEX=2;CharacterSet=936;DATABASE=R:\Sales Simulation\Simulation\Import\2007_07_19 @01 55;TABLE=Import#Csv

End Function

Sub TblCrt_FmLnkCsv__Tst()
Dim cFfnCsv$, cNmtNew$
'cFfnCsv$ = "R:\Sales Simulation\Simulation\Import\2007_07_19 @01 55\Import.Csv"
'cNmtNew$ = "tmpImp_Import"
'CrtTbl_FmLnkCsv_Tst = CrtTbl_FmLnkCsv(cFfnCsv, cNmtNew)
cFfnCsv$ = "R:\Sales Simulation\Simulation\Import\2007_07_19 @01 55\DataTotalEuro S01 BrandGp03-Nam\Val.csv"
cNmtNew$ = "tmpImp_Val"
TblCrt_FmLnkCsv cFfnCsv, cNmtNew
End Sub

Function TblCrt_FmLnkLdb(pFbLdb$, pLoadInstId&, pNDb$, pLnt$) As Boolean
Const cSub$ = "TblCrt_FmLnkLdb"
'Aim: Create a list of table in {pLnt} by referring {pFbLdb} & {pLoadInstId}
Dim Db As database: If Opn_Db_R(Db, pFbLdb) Then ss.A 1: GoTo E
Dim mLn_wQuote$: If Q_Ln(mLn_wQuote, pLnt) Then ss.A 2: GoTo E
Dim mSql$: mSql = "Select" & _
" [SdirHom] & 'Db' & Format([DbSno],'000') & '.Db' AS xFbTar," & _
" [NmHost] & '_' & [NDb] & '_' & [Nmt]                 AS xNmt" & _
" from tblLdbHdr h inner join tblLdbDet d on h.LoadInstId=d.LoadInstId where h.LoadInstId=" & pLoadInstId & " and Nmt in (" & mLn_wQuote & ")"
With Db.OpenRecordset(mSql)
    While Not .EOF
        Dim mNmt$:     mNmt = !xNmt
        Dim mFbTar$:   mFbTar = !xFbTar

        If TblCrt_FmLnkNmt(mFbTar, mNmt) Then ss.A 3: GoTo E
        .MoveNext
    Wend
    .Close
End With
GoTo X
R: ss.R
E:
X:
    Cls_Db Db
End Function

Function TblCrt_FmLnkLdb__Tst()
Debug.Print TblCrt_FmLnkLdb("M:\07 ARCollection\ARCollection\WorkingDir\PgmObj\modLdDb", 7, "RBPCSF", "IIM,IIC")
End Function

Function TblCrt_FmLnkLnt(SrcFb$, pLnt$, Optional pLntNew$ = "", Optional pInDb As database) As Boolean
'Aim: Create NonBlank({pLntNew},{pLnt}) in {pInDb} by linking {SrcFb}!{pLnt}
Const cSub$ = "TblCrt_FmLnkLnt"
On Error GoTo R
Dim mAnt$():      If Brk_Ln2Ay(mAnt, pLnt) Then ss.A 1: GoTo E
Dim mAntNew$():   If Brk_Ln2Ay(mAntNew, Fct.NonBlank(pLntNew, pLnt)) Then ss.A 2: GoTo E
Dim N%: N = Sz(mAnt)
Dim J%
For J = 0 To N - 1
    If TblCrt_FmLnkNmt(SrcFb, mAnt(J), mAntNew(J), pInDb) Then ss.A 3: GoTo E
Next
Exit Function
R: ss.R
E:
End Function

Function TblCrt_FmLnkLnt__Tst()
Const cSub$ = "TblCrt_FmLnkLnt_Tst"
Dim mLnt$, mFbSrc$, mLntNew$
Dim mResult As Boolean, mCase As Byte: mCase = 1
Select Case mCase
Case 1
    mLnt = "tblOdbcSql,tblFc"
    mFbSrc = "D:\SPLHalfWayHouse\MPSDetail\VerNew@2007_01_04\WorkingDir\PgmObj\MPS_RfhFc.Db"
    mLntNew = ""
End Select
mResult = TblCrt_FmLnkLnt(mFbSrc, mLnt, mLntNew)
End Function

Function TblCrt_FmLnkNmt(pFb$, T$, Optional TNew$ = "", Optional pInDb As database) As Boolean
'Aim: Create NonBlank({TNew},{T}) in {pInDb} by linking {pFb}!{T}
Const cSub$ = "TblCrt_FmLnkNmt"
Dim mNmt$: mNmt = NonBlank(TNew, T)
Dim mNmtSrc$: mNmtSrc = T
Dim mCnn$: mCnn = ";DATABASE=" & pFb ';DATABASE={pFb};TABLE={T}
TblCrt_FmLnkNmt = TblCrt_FmLnk(mNmt, mNmtSrc, mCnn, pInDb)
Exit Function
R: ss.R
E:
End Function

Function TblCrt_FmLnkNmt__Tst()
Dim mFb$: mFb = "c:\tmp\aa.Db"
Dim mNmt$: mNmt = "tmpLnk_AA"
Dim mNmtNew$: mNmtNew = "$AA"
Dim Db As database: If Crt_Db(Db, mFb, True) Then Stop
TblCrt_ByFldDclStr mNmt, "aa text 10, bb int", , , Db
If TblCrt_FmLnkNmt(mFb, mNmt, mNmtNew) Then Stop
End Function

Function TblCrt_FmLnkSetNmt__Tst()
Const cSub$ = "TblCrt_FmLnkLnt_Tst"
Dim mFbSrc$, mSetNmt$, mPfxNmt$
Dim mResult As Boolean, mCase As Byte: mCase = 1
Select Case mCase
Case 1
    mSetNmt = "tbl*"
    mFbSrc = "D:\SPLHalfWayHouse\MPSDetail\VerNew@2007_01_04\WorkingDir\PgmObj\MPS_RfhFc.Db"
    mPfxNmt = "$"
End Select
mResult = TblCrt_FmLnkLnt(mFbSrc, mSetNmt, mPfxNmt)
End Function

Function TblCrt_FmLnkSetWs(Pfx$, pSetWs$, Optional pPfxNmt$ = "", Optional pInDb As database) As Boolean
'Aim: Create table using pPfx + ws name in {pInDb} by linking {pFx}!{pSetWs}.
Const cSub$ = "TblCrt_FmLnkSetWs"
StsShw "Linking [" & Pfx & "]![" & pSetWs & "]......"
Dim mWb As Workbook: If Opn_Wb_R(mWb, Pfx) Then ss.A 1: GoTo E
Dim mAnWs$(): If Fnd_AnWs_BySetWs(mAnWs, mWb, pSetWs) Then ss.A 2: GoTo E
Dim mCnn$: mCnn = CnnStr_Xls(Pfx)
Dim J%
For J = 0 To Sz(mAnWs) - 1
    Dim mNmtSrc$: mNmtSrc = mAnWs(J) & "$"
    Dim mNmt$: mNmt = pPfxNmt & mAnWs(J)
    If TblCrt_FmLnk(mNmt, mNmtSrc, mCnn, pInDb) Then ss.A 3: GoTo E
Next
GoTo X
R: ss.R
E:
X:
    Clr_Sts
    Cls_Wb mWb, False, True
End Function

Function TblCrt_FmLnkWs(Pfx$, Optional pWsNm$ = "", Optional TNew$ = "", Optional pInDb As database) As Boolean
'Aim: Create NonBlank({TNew},{pWsNm}) in {pInDb} by linking {pFx}!{pWsNm}.  If {pWsNm} is not given, use FileName(pFx).
Const cSub$ = "TblCrt_FmLnkWs"
If pWsNm = "" Then pWsNm = Cut_Ext(Fct.Nam_FilNam(Pfx))
Dim mNmt$: mNmt = NonBlank(TNew, pWsNm)
Dim mCnn$: mCnn = CnnStr_Xls(Pfx)
Dim mNmtSrc$: mNmtSrc = pWsNm & "$"
TblCrt_FmLnkWs = TblCrt_FmLnk(mNmt, mNmtSrc, mCnn, pInDb)
Exit Function
R: ss.R
E:
'Excel 8.0;HDR=YES;IMEX=2;DATABASE=D:\Data\MyDoc\Development\ISS\Imports\PO\PUR904 (On-Line).xls;TABLE='PUR904 (On-Line)'
End Function

Function TblCrt_FmLnkWs__Tst()
Const cSub$ = "TblCrt_FmLnkWs_Tst"
Const cFx$ = "c:\tmp\aa.xls"
Const cWsNm$ = "aa"
Dim mWb As Workbook: If Crt_Wb(mWb, cFx, True, cWsNm) Then Stop
Dim mWs As Worksheet: Set mWs = mWb.Sheets(1)
If Set_Ws_ByLpAp(mWs, 1, "abc,def,xyz", 1, "a123", Now) Then Stop
If Cls_Wb(mWb, True) Then Stop
If TblCrt_FmLnkWs(cFx, cWsNm) Then Stop
End Function

Function TblCrt_FmLnkXls(Pfx$, Optional pPfx$ = "", Optional A As database) As Boolean
'Aim: Link all worksheets in {pFx} as tables in {A}
Const cSub$ = "TblCrt_FmLnkXls"
StsShw "Create tables by linking [" & Pfx & "]...."
Dim AnWs$():  If Fnd_AnWs(AnWs, Pfx) Then ss.A 1: GoTo E
Dim iWsNm, mA$
For Each iWsNm In AnWs
    Dim mWsNm$: mWsNm = iWsNm
    If TblCrt_FmLnkWs(Pfx, mWsNm, pPfx & mWsNm, A) Then mA = Add_Str(mA, mWsNm)
Next
If Len(mA) <> 0 Then ss.A 1, "Some ws {mA} in xls file cannot be linked", "mA", mA: GoTo E
GoTo X
R: ss.R
E:
X:
    Clr_Sts
End Function

Function TblCrt_FmLnkXls__Tst()
MsgBox TblCrt_FmLnkXls("c:\temp\LT\LT.xls")
End Function

Function TblCrt_FmMgeNRec_To1Fld(T, Optional pSepChr$ = CtComma, Optional pFillDta = False) As Boolean
'Aim: Create a table of name {T}_Lst of 2 fields from the first 2 fields of {T}.
'     The fields name of {T}_Lst is same as the first 2 fields of {T} with prefix [_Lst] in 2nd field
'     The 2nd field of {T}_Lst is always memo no matter what field type of 2nd field of {T}
'     The 1st field of {T}_Lst will the PrimaryKey and this PrimaryKey will be created.
'     Create empty {T}_Lst if pFillDta is false
Const cSub$ = "TblCrt_FmMgeNRec_To1Fld"
Dim mF1$: mF1 = CurrentDb.TableDefs(T).Fields(0).Name
Dim mF2$: mF2 = CurrentDb.TableDefs(T).Fields(1).Name
Dim mSql$
mSql = Fmt_Str("Select {0} into {1}_Lst from {1} where false", mF1, T)
If Run_Sql(mSql) Then ss.A 1: GoTo E
mSql = Fmt_Str("Alter table {0}_Lst Add COLUMN {1}_Lst Memo", T, mF2)
If Run_Sql(mSql) Then ss.A 2: GoTo E
If Not pFillDta Then Exit Function
Dim mLasF1$, mF2Lst$
With CurrentDb.OpenRecordset(Fmt_Str("Select {0},{1} from {2} order by {0},{1}", mF1, mF2, T))
    If .AbsolutePosition <> -1 Then mLasF1 = .Fields(0).Value
    While Not .EOF
        If mLasF1 = .Fields(0).Value Then
            mF2Lst = Add_Str(mF2Lst, CStr(Nz(.Fields(1).Value, "")), pSepChr)
        Else
            mSql = Fmt_Str("Insert into {0}_Lst ({1},{2}_Lst) values ('{3}','{4}')", T, mF1, mF2, mLasF1, mF2Lst)
            If Run_Sql(mSql) Then ss.A 3: GoTo E
            mLasF1 = .Fields(0).Value
            mF2Lst = .Fields(1).Value
        End If
        .MoveNext
    Wend
    mSql = Fmt_Str("Insert into {0}_Lst ({1},{2}_Lst) values ('{3}','{4}')", T, mF1, mF2, mLasF1, mF2Lst)
    If Run_Sql(mSql) Then ss.A 3: GoTo E
    .Close
End With
Exit Function
R: ss.R
E:
End Function

Sub TblCrt_FmMgeNRec_To1Fld__Tst()
'tmpMPS_SKUFacParam is from MPSDetail.Db
Const cNmt$ = "tmpMPS_SKUFacParam"
Const cNmt_x$ = "tmpMPS_SKUFacParam_x"
Const cSub$ = "TblCrt_FmMgeNRec_To1Fld_Tst"
DoCmd.CopyObject , cNmt_x, acTable, cNmt
Dim mSql$
mSql = Fmt_Str("Update {0} set SKU_FacParam=Fac & ': ' & SKU_FacParam", cNmt_x)
If Run_Sql(mSql) Then ss.A 1: GoTo E
mSql = Fmt_Str("Alter table {0} Drop Column Fac", cNmt_x)
If Run_Sql(mSql) Then ss.A 1: GoTo E

Dim mRslt As Boolean: mRslt = TblCrt_FmMgeNRec_To1Fld(cNmt_x, vbCrLf)
DoCmd.OpenTable cNmt_x & "_Lst"
Exit Sub
R: ss.R
E:
End Sub

Function TblCrt_ForEdtTbl(Qry_or_Tbl_NmSrc$, NPk As Byte, Optional TarTn$ = "", Optional pStructOnly As Boolean = False) As Boolean
'Aim: Create table {mNmtTar} from {Qry_or_Tbl_NmSrc}.  {mNmtTar}'s content comes from {Qry_or_Tbl_NmSrc}.
'{mNmTar} fmt: first {NPk} is same as {Qry_or_Tbl_NmSrc}, then a field [Change], then list of pair fields [xx] and [New xx]
Const cSub$ = "TblCrt_ForEdtTbl"
If NPk = 0 Then ss.A 1, "NPk must > 0", , "Qry_or_Tbl_NmSrc,mNmTar", Qry_or_Tbl_NmSrc, TarTn: GoTo E
Dim mNmTar$: mNmTar = NonBlank(TarTn, "tmpEdt_" & Qry_or_Tbl_NmSrc)
If Dlt_Tbl(mNmTar) Then ss.A 1: GoTo E

Dim mLnFld$
If IsTbl(Qry_or_Tbl_NmSrc) Then
    mLnFld = ToStr_Flds(CurrentDb.TableDefs(Qry_or_Tbl_NmSrc).Fields)
ElseIf IsQry(Qry_or_Tbl_NmSrc) Then
    mLnFld = ToStr_Flds(CurrentDb.QueryDefs(Qry_or_Tbl_NmSrc).Fields)
Else
    ss.A 1, "Given Qry_or_Tbl_NmSrc is not table or query": GoTo E
End If
Dim mAnFld$(): mAnFld = Split(mLnFld, CtComma)
Dim A$: A = mAnFld(0)
Dim J%: For J = 1 To NPk - 1
    A = ", " & mAnFld(J)
Next
A = A & ", " & "'' AS Changed"
Dim B$
For J = NPk To UBound(mAnFld)
    A = A & ", [" & mAnFld(J) & "],'' as [New " & mAnFld(J) & "]"
    B = Add_Str(B, "[New " & mAnFld(J) & "]=Null")
Next
A = A & ", '' As [Error During Import]"
Dim mSql$
mSql = Fmt_Str("Select {0} into {1} from {2}", A, mNmTar, Qry_or_Tbl_NmSrc)
If pStructOnly Then
    If Run_Sql(mSql & " Where False") Then ss.A 2: GoTo E
    Exit Function
End If
If Run_Sql(mSql) Then ss.A 3: GoTo E
mSql = Fmt_Str("Update {0} set {1}", mNmTar, B)
If Run_Sql(mSql) Then ss.A 4: GoTo E
Exit Function
R: ss.R
E:
End Function

Function TblCrt_ForEdtTbl__Tst()
Const cSub$ = "TblCrt_ForEdtTbl_Tst"
Dim mNmtqSrc$, mNmtTar$
Dim mRslt As Boolean, mCase As Byte: mCase = 2
Select Case mCase
Case 1
    mNmtqSrc = "tblUsr"
    mNmtTar = ""
Case 2
    mNmtqSrc = "tblCus"
    mNmtTar = ""
End Select
mRslt = TblCrt_ForEdtTbl(mNmtqSrc, 1, mNmtTar)
End Function

Function TblCrt_ParChd(TarTn$, TSrc$, pPar$, pChd$, Optional A As database) As Boolean
'Aim: Build TarTn of structure: Sno, Par, Chd, Lvl, from {TSrc} & {TNm}
'     Assume Struct: TarTn: {pPar}, {pChd}
Const cSub$ = "TblCrt_ParChd"
Dim mNmtSrc$: mNmtSrc = Rmv_SqBkt(TSrc)
If Chk_Struct_Tbl_SubSet(mNmtSrc, pPar & "," & pChd) Then ss.A 1: GoTo E
Dim mAyRoot&(): If Fnd_AyRoot(mAyRoot, TSrc, pPar, pChd) Then ss.A 3: GoTo E
Dim mNmtTar$: mNmtTar = Rmv_SqBkt(TarTn)
TblCrt_ByFldDclStr mNmtTar, "Sno Long, Par Long, Chd Long, Lvl Byte", 1, , A
Dim mRsTar As DAO.Recordset: If Opn_Rs(mRsTar, "Select * from [" & mNmtTar & "]") Then ss.A 5: GoTo E
Dim mSno&, mLvl As Byte: mSno = 0: mLvl = 0
Dim J%
Dim mAyPth&(), N%: N = Sz(mAyRoot)
For J = 0 To N - 1
    If J Mod 50 = 0 Then StsShw J & "(" & N & ") ..."
    TblCrt_ParChd_OneRec mRsTar, 0, mAyRoot(J), mLvl
    If TblCrt_ParChd_OneRoot(mAyRoot(J), mAyPth, mLvl, mRsTar, mNmtSrc, pPar, pChd) Then ss.A 6: GoTo E
Next
GoTo X
E:
X: RsCls mRsTar
End Function

Function TblCrt_ParChd__Tst()
'If TblCrt_FmLnkLnt("p:\workingdir\MetaDb.Db", "$Tbl,$TblR") Then Stop: GoTo E
'Dim mFx$, mWb1 As Workbook, mWb2 As Workbook, mWs As Worksheet
'If True Then
'    mFx = "c:\tmp\aa.xls"
'    If True Then
'        If TblCrt_FmLnkLnt("P:\WorkingDir\MetaAll.Db", "$Tbl,$TblR") Then Stop: GoTo E
'        If Run_Qry("qryTstCrtTblParChd") Then Stop: GoTo E
'        If Exp_SetNmtq2Xls("[#]Lst", mFx, True) Then Stop: GoTo E
'    End If
'    If Opn_Wb_RW(mWb1, mFx) Then Stop: GoTo E
'    Set mWs = mWb1.Sheets(1)
'    If WsFmtOL_ByCol(mWs.Range("A2"), 5, 6) Then Stop: GoTo E
'    mWb1.Save
'    mWb1.Application.Visible = True
'    Stop
'End If
'If True Then
'    mFx = "c:\tmp\bb.xls"
'    If TblCrt_ParChd("#Tmp", "$TblR", "TblTo", "Tbl") Then Stop: GoTo E
'
'    If Run_Sql("Alter table [#Tmp] Add NmPar Text(50), L Long, NmChd Text(50)") Then Stop: GoTo E
'    If Run_Sql("Update [#Tmp] m inner join [$Tbl] s" & _
'        " On m.Par=s.Tbl" & _
'        " Set m.NmPar=s.NmTbl" & _
'        " Where Par<>0") Then Stop: GoTo E
'    If Run_Sql("Update [#Tmp] set NmPar='Root' where Par=0") Then Stop: GoTo E
'    If Run_Sql("Update [#Tmp] set L=Lvl+1") Then Stop: GoTo E
'    If Run_Sql("Alter Table [#Tmp] Drop Column Lvl") Then Stop: GoTo E
'    If Run_Sql("Update [#Tmp] m inner join [$Tbl] s" & _
'        " On m.Chd=s.Tbl" & _
'        " Set m.NmChd=s.NmTbl") Then Stop: GoTo E
'
'    If Exp_SetNmtq2Xls("[#]Tmp", mFx, True) Then Stop: GoTo E
'    If Opn_Wb_RW(mWb2, mFx) Then Stop: GoTo E
'    Set mWs = mWb2.Sheets(1)
'    If WsFmtOL_ByCol(mWs.Range("A2"), 5, 6) Then Stop: GoTo E
'    mWb2.Save
'End If
'mWs.Application.Visible = True
'Stop
'GoTo X
'Exit Function
'E:
'X:
'    Cls_Wb mWb1, False, True
'    Cls_Wb mWb2, False, True
End Function

Function TblCrt_tmpXXX_Prm_By_qryOdbcXXX_0(QryNmsns$, Optional pLm$ = "") As Boolean
Const cSub$ = "Crt_"
Dim mNmPrm$: mNmPrm = "tmpOdbc" & QryNmsns & "_Prm"
Dim mAnq$(): If Fnd_Anq_ByPfx(mAnq, "qryOdbc" & QryNmsns & "_0") Then ss.A 3: GoTo E
If Run_Qry_ByAnq(mAnq, pLm) Then ss.A 4: GoTo E
If Not IsTbl(mNmPrm) Then ss.A 1, "Table mNmPrm not exist", eRunTimErr, "mNmPrm", mNmPrm: GoTo E
Exit Function
R: ss.R
E:
End Function

Function TblCrtFmCsv(pFfnCsv$, Optional TNew$ = "", Optional pAcs As Access.Application = Nothing) As Boolean
Const cSub$ = "TblCrtFmCsv"
On Error GoTo R
Dim mAcs As Access.Application: Set mAcs = Cv_Acs(pAcs)
Dim Db As database: Set Db = mAcs.CurrentDb
Dim mNmtNew$: If TNew = "" Then mNmtNew = Fct.Nam_FilNam(pFfnCsv) Else mNmtNew = TNew
Dlt_Tbl mNmtNew, Db
mAcs.DoCmd.TransferText acImportDelim, , mNmtNew, pFfnCsv, True
GoTo X
R: ss.R
E:
X: Set Db = Nothing
End Function

Function TblCrtFmCsv__Tst()
Dim mFfnCsv$
Dim mCase As Byte: mCase = 1
Select Case mCase
Case 1: mFfnCsv = "c:\Tmp\CsvChgTbl_20080518_175348(4).csv"
End Select
If TblCrtFmCsv(mFfnCsv, ">ChgTbl") Then Stop
DoCmd.OpenTable ">ChgTbl"
End Function

Function TblCrtFmLnkTxt(T, pFt$, SpecNm$, Optional pInDb As database) As Boolean
'Aim: Create {T} in {pInDb} by linking {pFt} using {SpecNm}.
Const cSub$ = "TblCrtFmLnkTxt"
If VBA.Dir(pFt) = "" Then ss.A 1, "Given txt file not found": GoTo E
'Text;DSN=A1 Link Specification;FMT=Fixed;HDR=NO;IMEX=2;CharacterSet=20127;DATABASE=C:\;TABLE=a1#txt
Dim mDir$: mDir = Fct.Nam_DirNam(pFt)
Dim mNmtSrc$:  mNmtSrc = Fct.Nam_FilNam(pFt)
Dim mFn$: mFn = Replace(mNmtSrc, ".", "#")
Dim mCnn$: mCnn = Fmt_Str("Text;DSN={0};FMT=Fixed;HDR=NO;IMEX=2;CharacterSet=20127;DATABASE={1};TABLE={2}", SpecNm, mDir, mFn)
TblCrtFmLnkTxt = TblCrt_FmLnk(T, mNmtSrc, mCnn, pInDb)
Exit Function
R: ss.R
E:
End Function

Sub TblCrtFmLnkTxt__Tst()
Dim mNmt$: mNmt = "A1"
Dim mFb$: mFb = "c:\aa.Db"
Dim mFt$: mFt = "c:\a1.txt"
Dim mNmSpec$: mNmSpec = "A1"
Dim mTxtSpec$: mTxtSpec = "I=Int3, AA=Txt10, B=Txt2, C=Txt3"
Dim Db As database:: If Crt_Db(Db, mFb, True) Then Stop
If Dlt_Tbl(mNmt, Db) Then Stop
If Dlt_TxtSpec(mNmSpec, Db) Then Stop
TxtSpecCrt_Fix mNmSpec, mTxtSpec, Db
If Dlt_Fil(mFt) Then Stop
Open mFt For Output As #1
Close #1
If TblCrtFmLnkTxt(mNmt, mFt, mNmSpec, Db) Then Stop
End Sub

Sub TblCrtFmTblF(Optional TTblF$ = "#TblF", Optional A As database)
'Aim: Create all tables as defined in {TTblF} to the Fb
'     #TblF: Pth,NmDb,NPk,StopAutoInc,TblAtr,NmTbl,SnoTblF,NmFld,DaoTy,FldLen,FmtTxt,IsReq,IsAlwZerLen,VdtTxt,VdtRul,DftVal
Const cSub$ = "TblCrtFmTblF"
On Error GoTo R
Dim mAyFld() As DAO.Field
If Chk_Struct_Tbl(TTblF, "Pth,NmDb,NPk,StopAutoInc,TblAtr,NmTbl,SnoTblF,NmFld,DaoTy,FldLen,FmtTxt,IsReq,IsAlwZerLen,VdtTxt,VdtRul,DftVal") Then ss.A 1: GoTo E
Dim mNmtTblF$: mNmtTblF = Q_S(TTblF, "[]")
Dim mAyFb$(): mAyFb = SqlSy("Select Distinct Pth & NmDb from " & mNmtTblF)
Dim iFb%
For iFb = 0 To Sz(mAyFb) - 1
    Dim Db As database: If Opn_Db_RW(Db, mAyFb(iFb)) Then ss.A 2: GoTo E
    Dim mAyNPk() As Byte, mAyStopAutoInc() As Boolean, mAyTblAtr&(), mAnt$()
    Dim mSql$: mSql = Fmt_Str("Select Distinct NPk,StopAutoInc,TblAtr,NmTbl from {0} where Pth & NmDb='{1}' order by NmTbl", mNmtTblF, mAyFb(iFb))
    If Fnd_LoAyV_FmSql(mSql, "NPk,StopAutoInc,TblAtr,NmTbl", mAyNPk, mAyStopAutoInc, mAyTblAtr, mAnt) Then ss.A 2: GoTo E
    Dim iNmt%
    For iNmt = 0 To Sz(mAnt) - 1
        StsShw "Creating Table " & mAnt(iNmt) & " ..."
        mSql = Fmt_Str("Select NmFld,DaoTy,FldLen,FmtTxt,IsReq,IsAlwZerLen,VdtTxt,VdtRul,DftVal from {0} where NmTbl='{1}' order by SnoTblF", mNmtTblF, mAnt(iNmt))
        Dim mRs As DAO.Recordset: If Opn_Rs(mRs, mSql) Then ss.A 3: GoTo E
        Dim J%: J = 0
        With mRs
            While Not .EOF
                ReDim Preserve mAyFld(J)
                If FldNew_FmRsTblF(mAyFld(J), mRs) Then ss.A 4: GoTo E
                J = J + 1
                .MoveNext
            Wend
            .Close
        End With
        TblCrt_ByFldAy mAnt(iNmt), mAyFld, mAyNPk(iNmt), mAyTblAtr(iNmt), mAyStopAutoInc(iNmt), Db

        'Add ZerRec
        If mAyNPk(iNmt) = 1 Then
            If Left(mAnt(iNmt), 3) <> "$Ty" Or mAnt(iNmt) = "$TypDta" Then
                Dim mAnFld$(): If Fnd_AnFld_ReqTxt(mAnFld, mAnt(iNmt), Db) Then ss.A 3: GoTo E
                Dim mLnFld$, mLnVal$
                mLnFld = "": mLnVal = ""
                Dim I%, NFld%: NFld = Sz(mAnFld)
                For I = 0 To NFld - 1
                    mLnFld = mLnFld & "," & mAnFld(I)
                    mLnVal = mLnVal & ",'-'"
                Next
                mSql = Fmt_Str("Insert into [${0}] ({0}{1}) values (0{2})", Mid(mAnt(iNmt), 2), mLnFld, mLnVal)
                If Run_Sql_ByDbExec(mSql, Db) Then ss.A 4: GoTo E
            End If
        End If
    Next
    Db.Close
Next
GoTo X
R: ss.R
E:
X:
    Cls_Db Db
    RsCls mRs
    Clr_Sts
End Sub

Sub TblCrtFmTblF__Tst()
TblCrt_FmLnkNmt "p:\workingdir\pgmobj\JMtcDb.Db", "#TblF"
TblCrtFmTblF
End Sub

Function TblCrtPk(T$, FnStr$, Optional A As database) As Boolean
'Aim: Create PrimaryKey on {T} by {FnStr}
Const cSub$ = "TblCrtPk"
On Error GoTo R
If Dlt_Idx(T, "PrimaryKey", A) Then ss.A 1: GoTo E
Dim mSql$: mSql = Fmt_Str("Create Index PrimaryKey on {0} ({1}) Primary", Q_SqBkt(T), FnStr)
If Run_Sql_ByDbExec(mSql, A) Then ss.A 2: GoTo E
Exit Function
R: ss.R
E:
End Function

Sub TblCrtSubDtaSheet(MstTn$, ChdTn$, MstFnStr$, Optional ChdFnStr$, Optional A As database)
Dim O As TableDef: Set O = DbNz(A).TableDefs(MstTn)
Dim OMst$
    OMst = AyJnComma(NmstrBrk(MstFnStr))
Dim OChd$
    OChd = IIf(ChdFnStr = "", MstFnStr, ChdFnStr)
    OChd = AyJnComma(NmstrBrk(OChd))
TblSetPrp O, "SubdatasheetName", ChdTn
TblSetPrp O, "LinkChildFields", OChd
TblSetPrp O, "LinkMasterFields", OMst
End Sub

Function TblCrtSubDtaSheet__Tst()
TblCrtSubDtaSheet "qryARInq_1_LvlAsOf", "qryARInq_1_LvlCus", "InstId"
End Function

Sub TblDrpPrp(T As TableDef, PrpNm$)
PrpDrp PrpNm, T.Properties
End Sub

Sub TblSetPrp(T As TableDef, PrpNm$, V)
If VarIsBlank(V) Then
    TblDrpPrp T, PrpNm
    Exit Sub
End If

If TblHasPrp(T, PrpNm) Then
    T.Properties(PrpNm) = V
Else
    T.Properties.Append T.CreateProperty(PrpNm, VarType(V), V)
End If
End Sub

Private Function TblCrt_ParChd_OneRec(pRsTar, pPar&, pChd&, pLvl As Byte) As Boolean
'     pRsTar: Sno, Par, Chd, Lvl
With pRsTar
    .AddNew
    !Par = pPar
    !Chd = pChd
    !Lvl = pLvl
    .Update
End With
End Function

Private Function TblCrt_ParChd_OneRoot(ByVal pRoot&, oAyPth&(), oLvl As Byte, pRsTar As DAO.Recordset, TSrc$, pPar$, pChd$) As Boolean
'Aim: Recursively write records to {pRsTar}.  Each root one extra write.
'     pRsTar: Sno, Par, Chd, Lvl
'     Assume TSrc has no []
oLvl = oLvl + 1

Dim mSql$: mSql = Fmt_Str("Select {0} from [{1}] where {2}={3} order by {0}", pChd, TSrc, pPar, pRoot)
Dim mAyId&(): mAyId = SqlIntoAy(mSql, mAyId)

Dim J%
For J = 0 To Sz(mAyId) - 1
    TblCrt_ParChd_OneRec pRsTar, pRoot, mAyId(J), oLvl

    Dim mIdx%: If Fnd_IdxLng(mIdx, oAyPth, mAyId(J)) Then ss.A 3:
    
    If mIdx < 0 Then
        Dim N%
        N = Sz(oAyPth)
        ReDim Preserve oAyPth(N): oAyPth(N) = mAyId(J)
        If TblCrt_ParChd_OneRoot(mAyId(J), oAyPth, oLvl, pRsTar, TSrc, pPar, pChd) Then ss.A 4:
        
        If N = 0 Then
            Clr_AyLng oAyPth
        Else
            ReDim Preserve oAyPth(N - 1)
        End If
    End If
Next

oLvl = oLvl - 1
End Function

