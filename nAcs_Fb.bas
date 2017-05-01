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

Sub FbRenToBackup(Fb$, Optional pKeepBackupLvl As Byte = 3)
If pKeepBackupLvl = 0 Then
    FfnDlt Fb
    Exit Sub
End If
If pKeepBackupLvl > 9 Then pKeepBackupLvl = 9
Dim mFfnn$: mFfnn = FfnCutExt(Fb)
Dim mExt$: mExt = FfnExt(Fb)
Dim mNxtFfnn$, mNxtBkNo As Byte: mNxtFfnn = FfnNxtBackup(mFfnn)
If mNxtBkNo >= 10 Or mNxtBkNo >= pKeepBackupLvl Then
    FfnDlt mNxtFfnn & mExt
    Name Fb As mNxtFfnn & mExt
    FfnSetRO mNxtFfnn & mExt
    Exit Sub
End If
If FfnIsExist(mNxtFfnn & mExt) Then FbRenToBackup mNxtFfnn & mExt, pKeepBackupLvl:     Exit Sub
Name Fb As mNxtFfnn & mExt
End Sub
