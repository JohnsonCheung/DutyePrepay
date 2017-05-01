Attribute VB_Name = "nAcs_Appa"
Option Compare Database
Option Explicit
Private X_Appa As Access.Application

Function Appa() As Access.Application
On Error GoTo X
Dim A$: A = X_Appa.Name
Set Appa = X_Appa
Exit Function
X: Set X_Appa = New Access.Application
Set Appa = X_Appa
End Function

Sub AppaAutoRun()
Dim A$: A = Environ("AppaAutoRun")
If A = "" Then Exit Sub
Run A
End Sub

Sub AppaClsAllTbl(Optional A As Access.Application)
Dim I As AccessObject
With AppaNz(A)
    For Each I In .CurrentData.AllTables
        With I
            If I.IsLoaded Then .DoCmd.Close acTable, .Name
        End With
    Next
End With
End Sub

Sub AppaClsCurDb(Optional A As Access.Application)
On Error Resume Next
AppaNz(A).CloseCurrentDatabase
End Sub

Sub AppaClsDb(Optional A As Access.Application)
On Error Resume Next
Dim App As Access.Application: Set App = AppaNz(A)
App.CloseCurrentDatabase
End Sub

Function AppaCpyObj(pLnObj_Tar$, pTypObj As AcObjectType, Optional pFb_Src$ = "", Optional pLnObj_Src$ = "") As Boolean
Const cSub$ = "Cpy_Obj"
Dim mAccess As Access.Application
If pFb_Src <> "" Then
    If Not IsFfn(pFb_Src) Then ss.A 1: GoTo E
    Set mAccess = G.gAcs
    If Opn_CurDb(mAccess, pFb_Src) Then Cls_CurDb mAccess: ss.A 1: GoTo E
End If

On Error GoTo R
Dim mAnObj_Tar$(): mAnObj_Tar = Split(pLnObj_Tar, CtComma)
Dim mAnObj_Src$(): mAnObj_Src = Split(Fct.NonBlank(pLnObj_Src, pLnObj_Tar), CtComma)
Dim N%: N = Sz(mAnObj_Src)
If Sz(mAnObj_Tar) <> N Then ss.A 1, "# of object names in Src & Tar are diff", , "Src,Tar", N, Sz(mAnObj_Tar): GoTo E
Dim J%
If pFb_Src = "" Then
    For J = 0 To N - 1
        DoCmd.CopyObject , mAnObj_Tar(J), pTypObj, mAnObj_Src(J)
    Next
Else
    For J = 0 To N - 1
        mAccess.DoCmd.CopyObject CurrentDb.Name, mAnObj_Tar(J), pTypObj, mAnObj_Src(J)
    Next
    Cls_CurDb mAccess
End If
Select Case pTypObj
Case Access.AcObjectType.acQuery: CurrentDb.QueryDefs.Refresh
Case Access.AcObjectType.acTable: CurrentDb.TableDefs.Refresh
End Select
Exit Function
R: ss.R
E:
X: If pFb_Src <> "" Then Cls_CurDb mAccess
End Function

Function AppaCpyObj__Tst()
Const cSub$ = "AppaCpyObjByPfx_Tst"
Dim mLnObj_Src$, mFb_Src$, mTypObj As Access.AcObjectType
Dim mResult As Boolean, mCase As Byte: mCase = 1
Select Case mCase
Case 1
    mLnObj_Src = "qryOdbcMPS_01_0_Prm,qryOdbcMPS_01_1_Fm_qEnv_qBrand"
    mFb_Src = "P:\MPSDetail\MPSDetail\MPS.Mdb"
    mTypObj = acQuery
End Select
mResult = AppaCpyObj(mLnObj_Src, mTypObj, mFb_Src)
End Function

Function AppaCpyObjByPfx(pPfx_Tar$, pTypObj As AcObjectType, Optional pFb_Src$ = "", Optional pPfx_Src$ = "") As Boolean
Const cSub$ = "AppaCpyObjByPfx"

If pFb_Src = "" And pPfx_Src = "" Then ss.A 1, "Cannot both pFb_Src & pPfx_Src be blank", , "pPfx_Tar,pTypObj", pPfx_Tar, ToStr_TypObj(pTypObj): GoTo E

Dim mAyTar$(): If Fnd_AnObj_ByPfx_InMdb(mAyTar$, pFb_Src, pPfx_Tar, pTypObj) Then ss.A 2: GoTo E
Dim mAySrc$(): If Repl_Pfx_InAy(mAySrc, pPfx_Src, mAyTar, pPfx_Tar) Then ss.A 2: GoTo E
Dim N%: N = Sz(mAySrc)
If Sz(mAyTar) <> N Then ss.A 1, "# of object names in Src & Tar are diff", , "Src,Tar", N, Sz(mAyTar): GoTo E

Dim mAccess As Access.Application
If pFb_Src <> "" Then
    If Not IsFfn(pFb_Src) Then ss.A 1: GoTo E
    Set mAccess = New Access.Application
    If Opn_CurDb(mAccess, pFb_Src) Then Cls_CurDb mAccess:  ss.A 1: GoTo E
End If
On Error GoTo R

Dim J%
If pFb_Src = "" Then
    For J = 0 To N - 1
        DoCmd.CopyObject , mAyTar(J), pTypObj, mAySrc(J)
    Next
Else
    For J = 0 To N - 1
        mAccess.DoCmd.CopyObject CurrentDb.Name, mAyTar(J), pTypObj, mAySrc(J)
    Next
    Cls_CurDb mAccess
    mAccess.Quit
    Set mAccess = Nothing
End If
Select Case pTypObj
Case Access.AcObjectType.acQuery: CurrentDb.QueryDefs.Refresh
Case Access.AcObjectType.acTable: CurrentDb.TableDefs.Refresh
End Select
Exit Function
R: ss.R
E:
X: If pFb_Src <> "" Then Cls_CurDb mAccess: mAccess.Quit: Set mAccess = Nothing
End Function

Function AppaCpyObjByPfx__Tst()
Const cSub$ = "AppaCpyObjByPfx_Tst"
Dim mPfx_Src$, mFb_Src$, mTypObj As Access.AcObjectType
Dim mResult As Boolean, mCase As Byte: mCase = 1
Select Case mCase
Case 1
    mPfx_Src = "qryOdbcFc_0"
    mFb_Src = "P:\MPSDetail\MPSDetail\WorkingDir\PgmObj\RfhFc.Mdb"
    mTypObj = acQuery
End Select
mResult = AppaCpyObjByPfx(mPfx_Src, mTypObj, mFb_Src)
End Function

Function AppaCrtFb(Fb$, Optional Locale$ = dbLangGeneral, Optional A As Access.Application) As Access.Application
Dim O As Access.Application: Set O = Appa
O.DBEngine.CreateDatabase(Fb, Locale).Close
O.OpenCurrentDatabase Fb
Set AppaCrtFb = O
End Function

Function AppaCrtPja(PjNm$, Optional Pth$, Optional A As Access.Application) As vbproject
Dim F$: F = PjNm_NewFmda(PjNm, Pth)
Set AppaCrtPja = FmdaCrt(F, A)
End Function

Function AppaDbAy(Optional A As Access.Application) As database()
Dim I As Workspace, D As database
Dim O() As database
For Each I In AppaNz(A).DBEngine.Workspaces
    For Each D In I.Databases
        PushObj O, D
    Next
Next
AppaDbAy = O
End Function

Function AppaDtaDb(Optional A As Access.Application) As database
Set AppaDtaDb = AppaNz(A).DBEngine.Workspaces(0).Databases(0)
End Function

Function AppaDtaFb$()
AppaDtaFb = FfnAddFnSfx(CurrentDb.Name, "_Data")
End Function

Sub AppaFmdaCrt__Tst()
Const N$ = "aaaa"
Dim F$: F = PjNm_NewFmda(N)
FfnDltIfExist F
Dim P As vbproject: Set P = FmdaCrt(F)
Debug.Assert P.FileName = F
Debug.Assert P.Name = N
Stop
End Sub

Function AppaNDb%(Optional P As Access.Application)
Dim I As Workspace
Dim O%
For Each I In AppaNz(P).DBEngine.Workspaces
    O = O + I.Databases.Count
Next
AppaNDb = O
End Function

Function AppaNWrkSpc%(Optional A As Access.Application)
AppaNWrkSpc = AppaNz(A).DBEngine.Workspaces.Count
End Function

Function AppaNz(A As Access.Application) As Access.Application
If IsNothing(A) Then
    Set AppaNz = Access.Application
Else
    Set AppaNz = A
End If
End Function

Function AppaOpnPj(Fmda$, Optional A As Access.Application) As vbproject
Dim App As Access.Application: Set App = AppaNz(A)
FfnAsstExt Fmda, ".mda", "OpnAppaPj"
App.OpenCurrentDatabase Fmda
Set AppaOpnPj = App.Vbe.VBProjects(1)
End Function

Function AppaOpnPjFmda(Fmda$, Optional A As Access.Application) As vbproject
Dim App As Access.Application: Set App = AppaNz(A)
AppaClsDb App
App.OpenCurrentDatabase Fmda
Set AppaOpnPjFmda = App.Vbe.VBProjects(1)
End Function

Sub AppaQuit()
On Error Resume Next
X_Appa.Quit
Set X_Appa = Nothing
End Sub

Sub AppaSavRec(Optional Appa As Access.Application)
AppaNz(Appa).DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
End Sub

Sub GoRec(Optional Where As AcRecord = acNext)
DoCmd.GoToRecord , , Where
End Sub
