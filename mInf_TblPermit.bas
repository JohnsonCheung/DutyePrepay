Attribute VB_Name = "mInf_TblPermit"
Option Compare Database
Option Explicit

Function TblPermitDate(PermitId&) As Date
Dim W$: W = FmtQQ("Permit=?", PermitId)
TblPermitDate = TblFldV("Permit", "Permit", W)
End Function

Function TblPermitIdByNo&(PermitNo$)
Dim W$: W = FmtQQ("PermitNo='?'", PermitNo)
TblPermitIdByNo = TblFldToLng("Permit", "Permit", W)
End Function

Sub TblPermitIdByNo__Tst()
Const A$ = "ND04VFOF00JBEN"
Debug.Assert TblPermitIdByNo(A) = 1692
End Sub

Function TblPermitNo$(PermitId&)
Dim W$: W = FmtQQ("Permit=?", PermitId)
TblPermitNo = TblFldV("Permit", "Permit", W)
End Function
