Attribute VB_Name = "nDao_Lnk"
Option Compare Database
Option Explicit
Const LnkTn = "LnkDfn"

Function LnkCnnSy(Optional A As database) As String()
Dim D As database: Set D = DbNz(A)
Dim T$()
T = OyPrp_Str(TblAy(A), "Connect")
LnkCnnSy = AyRmvBlank(T)
End Function

Sub LnkCnnSy__Tst()
AyBrw LnkCnnSy
End Sub

Sub LnkCrt(T$, Src$, CnnStr$, Optional A As database)
Dim D As database: Set D = DbNz(A)
TblDrp T, D
Dim Tbl As New DAO.TableDef
With Tbl
    .Connect = CnnStr
    .Name = T
    .SourceTableName = Src
    D.TableDefs.Append Tbl
End With
End Sub

Sub LnkFb(T$, Fb$, Optional SrcTn$, Optional A As database)
Dim Cnn$: Cnn = CnnStrLnkFb(Fb)
Dim S$: S = StrNz(SrcTn, T)
LnkCrt T, S, Cnn, A
'Excel 8.0;HDR=YES;IMEX=2;DATABASE=D:\Data\MyDoc\Development\ISS\Imports\PO\PUR904 (On-Line).xls;TABLE='PUR904 (On-Line)'
End Sub

Sub LnkFx(T$, Fx$, Optional SrcWsNm$, Optional A As database)
Dim W$: W = SrcWsNm: If W = "" Then W = FxFstWsNm(Fx)
Dim Cnn$: Cnn = CnnStrLnkFx(Fx)
Dim S$: S = W & "$"
LnkCrt T, S, Cnn, A
'Excel 8.0;HDR=YES;IMEX=2;DATABASE=D:\Data\MyDoc\Development\ISS\Imports\PO\PUR904 (On-Line).xls;TABLE='PUR904 (On-Line)'
End Sub

Sub LnkRfh()
Dim J%

Dim Dt As Dt
    Dt = TblDt(LnkTn)
    
Dim Fny$()
Dim DrAy
    With Dt
        Fny = .Fny
        DrAy = .DrAy
    End With

Dim IsEr_1 As Boolean
    Const A = ""
    Dim B$
        B = Join(Fny, " ")
    If A <> B Then IsEr_1 = True
If IsEr_1 Then Er "{C_LnkDfnTblNm} format should be {A}, But now it is {B}", LnkTn, A, B

Dim AppNy$()
Dim VerAy() As Byte
     DtAsgCol Dt, ApSy("AppNm", "Ver"), AppNy, VerAy
    Stop
Dim ErFbAy$()
    For J = 0 To UB(ErFbAy)
        If Not Fso.FileExists(ErFbAy(J)) Then
            Push ErFbAy, "File not exist[" & ErFbAy(J) & "]"
        End If
    Next

If Sz(ErFbAy) <> 0 Then
    'Er Msg
End If

Dim U&
For J% = 0 To U
    Dim Dr()
        Dr = DrAy(J)
    Dim AppNm$
    Dim TblNm$
    Dim Ver As Byte
    Dim Ext$
        AppNm = Dr(1)
        TblNm = Dr(0)
        Ver = Dr(0)
        Ext = Dr(0)
    Dim DbHasTblExist As Boolean
        
    'LnkDfn_Rfh TblNm, DbHasTblExist, B
Next
End Sub

Function LnkRfh__Tst()
If Crt_SessDta(1) Then Stop
'If LnkRfh(1) Then Stop
End Function

Function LnkRfh_ByRsLnkDef(pRsLnkDef As DAO.Recordset) As Boolean
Const cSub$ = "LnkRfh_ByRsLnkDef"
'Aim: Delete all link tables in CurrentDb
'     Create NonBlank(!NmtNew, !Nmt) in currentdb to link !Nmt in !Ffn of Type !NmLnkTyp
'     Assume pRsLnkDef has structure: Nmt, InFfn, NmLnkTyp, NmtNew
If Dlt_Tbl_ByLnk Then ss.A 1: GoTo E
On Error GoTo R
With pRsLnkDef
    While Not .EOF
        Select Case !NmTypLnk
        Case "XlsWs"
            If TblCrt_FmLnkWs(!InFfn, !Nmt, Nz(!NmtNew, "")) Then ss.A 2: GoTo E
        Case "MdbTbl"
            If TblCrt_FmLnkNmt(!InFfn, !Nmt, Nz(!NmtNew, "")) Then ss.A 3: GoTo E
        Case Else
        'Case "TxtFil"
            ss.A 4, "Unexpected Link Type", , "!Nmt,InFfn,!NmLnkTyp,!NmtNew", !Nmt, !InFfn, !NmLnkTyp, !NmtNew: GoTo E
        End Select
        .MoveNext
    Wend
End With
Exit Function
R: ss.R
E:
End Function

Sub LnkRfh1(pTrc&)
'Aim: Create link tables in CurDb for each record in "tblLnkTbl" & "tblLnkTblMdbSrc"
Const cSub$ = "LnkRfh"
On Error GoTo R
Dim mDirSess$: mDirSess = Fct.CurMdbDir & Format(pTrc, "00000000") & "\": If VBA.Dir(mDirSess, vbDirectory) = "" Then ss.A 1, , "[Sess Sub Dir] does not exist in currentDb", "CurDb", CurrentDb.Name: GoTo E
LnkRfh_Chk_tblLnkTbl
Dim mFb_modU$:  mFb_modU = Sdir_PgmObj & "mda"
Dim mFb_Dta$:   mFb_Dta = Sdir_Wrk & Fct.CurMdbNam & "_Data.mdb"
Dim xFfn$, xNmtSrc$, mSql$, mLnkLib
mSql = _
"Select      Nmt,LnkLib,FbSrc" & _
" from       tblLnkTbl_NewVer l" & _
" inner join tblLnkTblMdbSrc  s" & _
" on         l.MdbSrcId=s.MdbSrcId" & _
" order by   LnkLib"
With CurrentDb.OpenRecordset(mSql)
    While Not .EOF
        mLnkLib = Nz(!LnkLib, "")
        xNmtSrc = !Nmt
        If mLnkLib = "modU" Then
            xFfn = mFb_modU
        ElseIf mLnkLib = "" Then
            xFfn = mFb_Dta
        ElseIf Left(mLnkLib, 3) = "Tp:" Then
            'Nmt    LnkLib
            'aaa    Tp:TpNam!ssss
            Dim mA$: mA = Mid$(mLnkLib, 4)  'TpNam!ssss
            Dim mP%: mP = InStr(mA, "!")    'Pos of !
            Dim mTp$
            If mP > 0 Then
                mTp$ = Left(mA, mP - 1)     'TpNam
                xNmtSrc = Mid(mA, mP + 1)   'sss
            Else
                mTp$ = mA                   'TpNam
            End If
            If Fnd_Fn_By_Tp_n_CurFnn(mA, mTp, Fct.CurMdbNam) Then ss.A 1: GoTo E
            xFfn = mDirSess & mA
        ElseIf mLnkLib = "MdbSrc" Then
            xFfn = !FbSrc
        Else
            xFfn = Sdir_PgmObj & mLnkLib
        End If
        'StsShw "Linking [" & !Nmt & "] to [" & xFfn & "] ........"
        If TblCrt_FmLnkNmt(xFfn, xNmtSrc$, !Nmt) Then ss.A 1: GoTo E
        .MoveNext
    Wend
    .Close
End With
GoTo X
R: ss.R
E:
X: Clr_Sts
End Sub

Sub LnkTnBrw()
TblBrw LnkTn
End Sub

Sub LnkTnRfh(Optional A As database)
Dim ODt As Dt: ODt = DtSel(CnnDt, "TblNm AppNm Ver CnnStr")
LnkTnClr A
TblInsDt LnkTn, ODt, A
End Sub

Function LnkTny(Optional A As database) As String()
Dim B() As TableDef
B = OySelPrpNe(TblAy, "Connect", "")
LnkTny = OyPrp_Nm(B)
End Function

Sub LnkTny__Tst()
AyBrw LnkTny
End Sub

Private Function LnkRfh_Chk_tblLnkMdbSrc() As Boolean
'Aim: Check tblLnkTbl is in valid format
LnkRfh_Chk_tblLnkMdbSrc = Chk_Struct_Tbl("tblLnkMdbSrc", "Nmt,LnkLib,InUse,MdbSrcId")
End Function

Private Function LnkRfh_Chk_tblLnkTbl() As Boolean
'Aim: Check tblLnkTbl is in valid format
Const cSub$ = "LnkRfh_TblCrtLnkTbl"
On Error GoTo R
If Not Chk_Struct_Tbl("tblLnkTbl_NewVer", "Nmt,LnkLib,InUse,MdbSrcId") Then Exit Function
If Run_Sql("Create table tblLnkTbl_NewVer (Nmt Text(50), LnkLib Text(50), InUse YesNo, MdbSrcId Integer)") Then ss.A 1: GoTo E
Exit Function
R: ss.R
E:
End Function

Private Sub LnkTnClr(A As database)
DbRunSql "Delete From " & LnkTn, A
End Sub

Private Sub LnkTnCrt(A As database)
Dim S$: S = FmtQQ("Create Table ? (TblNm Text(20) PRIMARY KEY, AppNm Text(20), Ver Byte, CnnStr Memo)", LnkTn)
DbRunSql S, A
End Sub

Private Sub LnkTnDrp(A As database)
TblDrp LnkTn, A
End Sub

Private Sub LnkTnEns(A As database)
If Not LnkTnIsExist(A) Then LnkTnCrt A
End Sub

Private Function LnkTnIsExist(A As database) As Boolean
LnkTnIsExist = DbHasTbl(LnkTn, A)
End Function

