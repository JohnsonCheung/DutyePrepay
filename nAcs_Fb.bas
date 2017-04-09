Attribute VB_Name = "nAcs_Fb"
Option Compare Database
Option Explicit

Function FbAppa(Fb$) As Access.Application
Dim O As Access.Application
Set O = New Access.Application
O.OpenCurrentDatabase Fb
Set FbAppa = O
End Function

Sub FbCompact(Fb$, Optional BackupLvl% = 3)
Dim A$
    A = Fb & "_Compact.accdb"
    FfnDltIfExist A
    DAO.DBEngine.CompactDatabase Fb, A
FfnRenBackup Fb, BackupLvl
Name A As Fb
End Sub

Function FbCompact___Tst()
FbCompact ("M:\07 ARCollection\ARCollection\WorkingDir\ARCollection_Data.mdb")
End Function

Sub FbCrt(Fb$, Optional Locale$ = dbLangGeneral)
DAO.DBEngine.CreateDatabase Fb, Locale
End Sub

Function FbCur$()
FbCur = CurrentDb.Name
End Function

Function FbCurPth$()
FbCurPth = FfnPth(FbCur)
End Function

Function FbDb(Fb$) As database
Stop
End Function

Sub FbNew(Fb$, Optional Locale$ = dbLangGeneral)
DbNew(Fb, Locale).Close
End Sub

Function FbRenToBackup(pFb$, Optional pKeepBackupLvl As Byte = 3) As Boolean
Const cSub$ = "FbRenToBackup"
If pKeepBackupLvl = 0 Then
    If Dlt_Fil(pFb) Then ss.A 1: GoTo E
    Exit Function
End If
If pKeepBackupLvl > 9 Then pKeepBackupLvl = 9
Dim mFfnn$, mExt$: If Brk_Ffn_To2Seg(mFfnn, mExt, pFb) Then ss.A 1: GoTo E
Dim mNxtFfnn$, mNxtBkNo As Byte: Fnd_NxtBkFfnn mFfnn, mNxtFfnn, mNxtBkNo
If mNxtBkNo >= 10 Or mNxtBkNo >= pKeepBackupLvl Then
    If Dlt_Fil(mNxtFfnn & mExt, True) Then ss.A 1: GoTo E
    If Ren_Fil(pFb, mNxtFfnn & mExt) Then ss.A 2: GoTo E
    If Set_FilRO(mNxtFfnn & mExt) Then ss.A 3: GoTo E
    Exit Function
End If
If VBA.Dir(mNxtFfnn & mExt) <> "" Then If FbRenToBackup(mNxtFfnn & mExt, pKeepBackupLvl) Then Exit Function
If Ren_Fil(pFb, mNxtFfnn & mExt) Then ss.A 2: GoTo E
Exit Function
R: ss.R
E: FbRenToBackup = True: ss.B cSub, cMod, "pFb,pKeepBackupLvl", pKeepBackupLvl
End Function
