Attribute VB_Name = "mFrmSwitchboard_nV170302_DbFix"
Option Compare Database
Option Explicit

Sub DbFix()
TblPermit_AddCol_DteImp
End Sub

Sub DbFix__Tst()
DbFix
End Sub

Private Sub TblPermit_AddCol_DteImp()
Dim D As database
Set D = AppaDtaDb
TblAddFld "Permit", "DteImp", dbDate, D
TblAddFld "Permit", "IsCur", dbBoolean, D
TblAddFld "Permit", "CanImp", dbBoolean, D
D.Close
End Sub
