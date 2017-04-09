Attribute VB_Name = "nDta_nScl_Dr"
Option Compare Database
Option Explicit

Function DrNewScl(SemiColonLin$) As Variant()
Dim Ay$(): Ay = SclSy(SemiColonLin)
Dim O()
Dim J&
Dim U&: U = UB(Ay)
ReSz O, U
For J = 0 To U
    O(J) = VarSemiColonFldRev(Ay(J))
Next
DrNewScl = O
End Function

Function DrNewSclWithTy(SemiColonLin$, DaoTyAy() As DAO.DataTypeEnum)
Dim OTy() As VbVarType
Dim OAy$()
Dim OU&
Dim O()
    OTy = AyMapInto(DaoTyAy, OTy, "VbTyByDaoTy")
    OAy = SclSy(SemiColonLin)
    OU = UB(OAy)
    ReSz O, OU

Dim J&
Dim TU&:                  TU = UB(OTy)
Dim U&: U = Min(OU, TU)
For J = 0 To U
    O(J) = VarCv(OAy(J), OTy(J))
Next
For J = U + 1 To Max(OU, TU)
    O(J) = OAy(J)
Next
DrNewSclWithTy = O
End Function

Sub DrNewSclWithTy__Tst()
Dim L$
Dim T() As DataTypeEnum: T = DaoTySclTyAy("TXT;INT;LNG;DTE;YES")
Dim Act()
Dim Exp()
ReDim Exp(4)
Exp(0) = "1"
Exp(1) = 2
Exp(2) = CLng(3)
Exp(3) = #1/1/2017#
Exp(4) = True
L = "1;2;3;2017-1-1;true": Act = DrNewSclWithTy(L, T): AyAsstEqExa Exp, Act
End Sub

Function DrScl$(Dr)
Dim O$(): O = AyMapInto(Dr, O, "VarSemiColonFld")
DrScl = AyJn(O, ";")
End Function
