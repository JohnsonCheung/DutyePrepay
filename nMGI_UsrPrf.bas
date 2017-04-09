Attribute VB_Name = "nMGI_UsrPrf"
Option Explicit
Private X As UsrPrf
Private Const cLnItm$ = "Usr Dpt Fy Env Lvl Brand"
Type UsrPrf
    NmUsr As String
    NmDpt As String
    NmFy As String
    NmEnv As String
    NmLvl As String
    NmBrand As String
    Usr As Long
    Dpt As Long
    Fy As Long
    Env As Long
    Lvl As Long
    Brand As Long
End Type

Sub Gen_LetXXX()
Dim A$: A = ResStr("LetXXX", "nMGI_UsrPrf")
Debug.Print StrExpand(A, cLnItm, vbLf)
End Sub

Sub LetXXX()
'Property Let UsrPrf_{B}(p{B}&)
'Const cSub$ = "UsrPrf_{B}"
'If Run_Sql_ByDbExec(Fmt_Str("Update tblUsr SET {B}={B} Where Usr={1}", p{B}, X.UsrPrf_Usr), CodeDb) Then ss.A 1: GoTo E
'Dim mA$: If Fnd_ValFmSql(mA, Fmt_Str("Select Nm{B} from tbl{B} where {B}={B}", p{B}), CodeDb) Then ss.A 2: GoTo E
'X.Nm{B} = mA
'X.{B} = p{B}
'Exit Function
'E: ss.B cSub, cMod, "p{B}", p{B}
'End Property
End Sub

Sub UsrPrf_AsstPwd(NmUsr$, Pwd$)
Const cMsg$ = "Invalid User Id / Password"
If Pwd = "" Or NmUsr = "" Then GoTo X
Dim Rs As DAO.Recordset
Set Rs = CurrentDb.OpenRecordset("select * from tblUsr where NmUsr='" & NmUsr & CtSngQ)
With Rs
    If .AbsolutePosition = -1 Then GoTo X
    If Not !Enabled Then .Close: Er "User profile is disabled"
    If Pwd <> !Password.Value Then GoTo X
    '=====Login is OK
    X = RsUsrPrf(Rs)
    If UsrPrf_zzSetLoginToReg(!Usr.Value, True) Then Exit Sub
    .Edit
    !LoginCnt = !LoginCnt + 1
    !LasLoginDte = Now
    .Update
    .Close
End With
Exit Sub
X: Er "Invalid User Id / Password"
End Sub

Function UsrPrf_Login(NmUsr$) As Boolean
Dim Pwd$:
    Dim S$
    S = FmtQQ("Select password from tblUsr where NmUsr='?'", NmUsr)
    Pwd = SqlStr(S)
UsrPrf_AsstPwd NmUsr, Pwd
End Function

Function UsrPrf_Dpt&(): UsrPrf_Dpt = X.Dpt: End Function
Function UsrPrf_Fy&(): UsrPrf_Fy = X.Fy: End Function
Function UsrPrf_Env&(): UsrPrf_Env = X.Env: End Function
Function UsrPrf_Lvl&(): UsrPrf_Lvl = X.Lvl: End Function
Function UsrPrf_Brand&(): UsrPrf_Brand = X.Brand: End Function
Function UsrPrf_Login__Tst()
If UsrPrf_Login("Johnson") Then Stop
End Function
Function UsrPrf_NmBrand$(): UsrPrf_NmBrand = X.NmBrand: End Function
'Debug.Print StrExpand("Function Nm{B}$(): Nm{B} = X.{B}: End Function", "Usr,Dpt,Fy,Env,Lvl,Brand", vbLf)
Function UsrPrf_Usr&(): UsrPrf_Usr = X.Usr: End Function
Function UsrPrf_NmDpt$(): UsrPrf_NmDpt = X.NmDpt: End Function
Function UsrPrf_NmFy$(): UsrPrf_NmFy = X.NmFy: End Function
Function UsrPrf_NmEnv$(): UsrPrf_NmEnv = X.NmEnv: End Function
Function UsrPrf_NmLvl$(): UsrPrf_NmLvl = X.NmLvl: End Function
'Debug.Print StrExpand("Function Nm{B}$(): Nm{B} = X.{B}: End Function", "Usr,Dpt,Fy,Env,Lvl,Brand", vbLf)
Function UsrPrf_NmUsr$(): UsrPrf_NmUsr = X.NmUsr: End Function
Function UsrPrf_zzChkLogin() As Boolean
Const cSub$ = "zzChkLogin"
Dim App$: App = SysCfg_App
Dim A$: A = GetSetting(App, "UsrPrf", "HasLogin"): If A <> "True" Then UsrPrf_zzLoginAgain App: Exit Function
Dim AccessTim As Date: AccessTim = GetSetting(App, "UsrPrf", "AccessTim")
If DateDiff("h", AccessTim, Now()) > 1 Then UsrPrf_zzLoginAgain (App): Exit Function
If UsrPrf_Usr <= 0 Then
    Dim UsrId%: UsrId% = GetSetting(App, "UsrPrf", "Usr")
    If UsrId <= 0 Then Er "Cannot get user id"
    UsrPrf_zzGetUsrPrf_ByUsr UsrId
End If
SaveSetting App, "UsrPrf", "AccessTim", Now
End Function

Function UsrPrf_zzLoginAgain(App$) As Boolean
'Aim: Must login success (check against tblUsr), otherwise, quit.  If success, Reg: Usr & "AccessTim is set & X will be set
If SysCfg_IsNoLogin Then
    SaveSetting App, "UsrPrf", "AccessTim", Now
    'TblCrt_FmLnkNmt Sffn_Dta, "tblUsr"
    Dim Rs As DAO.Recordset
        Set Rs = CurrentDb.OpenRecordset("select * from tblUsr")    ' Assume user is not one record
    If Rs.AbsolutePosition <> -1 Then
        RsUsrPrf Rs
        SaveSetting App, "UsrPrf", "Usr", X.Usr
        Rs.Close
        Exit Function
    End If
    With Rs
        .AddNew
        !NmUsr = "NoLogin"
        !Password = "password"
        !UsrLvl = "T"
        .Update
        .Close
    End With
    Set Rs = CurrentDb.TableDefs("tblUsr").OpenRecordset
    X = RsUsrPrf(Rs)
    SaveSetting App, "UsrPrf", "Usr", X.Usr
    Rs.Close
    Exit Function
End If
If FrmOpn("frmLoginAgain", , True).Som Then Application.Quit
SaveSetting App, "UsrPrf", "AccessTim", Now
SaveSetting App, "UsrPrf", "Usr", X.Usr
End Function

Function UsrPrf_zzSetLoginToReg(UsrId%, pLoginOk As Boolean) As Boolean
Dim mApp$: mApp = SysCfg_App
SaveSetting mApp, "UsrPrf", "HasLogin", pLoginOk
SaveSetting mApp, "UsrPrf", "AccessTim", Now
SaveSetting mApp, "UsrPrf", "Usr", UsrId
End Function

Private Function RsUsrPrf(Rs As DAO.Recordset) As UsrPrf
Dim O As UsrPrf
With Rs
    On Error Resume Next
    O.Brand = Nz(!Brand, 0)
    O.Env = Nz(!Env, 0)
    O.Dpt = Nz(!Dpt, 0)
    O.Usr = Nz(!Usr, 0)
    O.Lvl = Nz(!Lvl, "")
    O.Fy = Nz(!Fy, Dte2FyNo)
End With
RsUsrPrf = O
End Function

Private Function UsrPrf_zzGetUsrPrf_ByUsr(UsrId%) As Boolean
TblCrt_FmLnkNmt Sffn_Dta, "tblUsr"
Dim Rs As DAO.Recordset
    Set Rs = CurrentDb.OpenRecordset("Select * from tblUsr where Usr=" & UsrId)
    RsUsrPrf Rs
    Rs.Close
End Function

Private Function XX() As UsrPrf
UsrPrf_zzChkLogin
XX = X
End Function
