Attribute VB_Name = "nDao_Lnk"
Option Compare Database
Option Explicit
Const LnkTn = "LnkDfn"

Sub LnkCrt(T$, S$, Cnn$, Optional A As database)
Dim D As database: Set D = DbNz(A)
TblDrp T, D
Dim Tbl As New DAO.TableDef
With Tbl
    .Connect = Cnn
    .Name = T
    .SourceTableName = S
    D.TableDefs.Append Tbl
End With
End Sub

Sub LnkCrt_Fx(T$, Fx$, Optional SrcWsNm$, Optional A As database)
Dim W$: W = SrcWsNm: If W = "" Then W = FxFstWsNm(Fx)
Dim Cnn$: Cnn = CnnStrFx(Fx)
Dim S$: S = W & "$"
LnkCrt T, S, Cnn, A
'Excel 8.0;HDR=YES;IMEX=2;DATABASE=D:\Data\MyDoc\Development\ISS\Imports\PO\PUR904 (On-Line).xls;TABLE='PUR904 (On-Line)'
End Sub

Sub LnkFbNew(T$, Fb$, Optional SrcTn$, Optional A As database)
Dim Cnn$: Cnn = CnnStrFb(Fb)
Dim S$: S = StrNz(SrcTn, T)
LnkCrt T, S, Cnn, A
'Excel 8.0;HDR=YES;IMEX=2;DATABASE=D:\Data\MyDoc\Development\ISS\Imports\PO\PUR904 (On-Line).xls;TABLE='PUR904 (On-Line)'
End Sub

Function LnkNy(Optional A As database) As String()
Dim D As database: Set D = DbNz(A)
Dim T() As TableDef: T = DbTblDefAy(A)
Dim T1$(): T1 = ObjAyPrp(T, "Connect", T1)
LnkNy = AyRmvBlank(T1)
End Function

Sub LnkNy__Tst()
AyBrw LnkNy
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

Sub LnkTnBrw()
TblBrw LnkTn
End Sub

Sub LnkTnRfh(Optional A As database)
Dim ODt As Dt: ODt = DtSel(CnnDt, "TblNm AppNm Ver CnnStr")
LnkTnClr A
TblInsDt LnkTn, ODt, A
End Sub

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
