Attribute VB_Name = "nDta_Fld"
Option Compare Database
Option Explicit
Type FldTySng
    Ty As DAO.DataTypeEnum
    F() As String
End Type
Type FldTyMul
    TyAy() As FldTySng
End Type

Function FldTyMulBrk(A$) As FldTyMul
Dim O As FldTyMul
Dim B$(): B = AyTrim(Split(A, "|"))
Dim U%: U = UB(B)
Dim OO() As FldTySng: ReDim OO(U)
Dim J%
For J = 0 To U
    OO(J) = FldTySngBrk(B(J))
Next
O.TyAy = OO
FldTyMulBrk = O
End Function

Sub FldTyMulBrk__Tst()
Dim A$: A = "TXT: A B [C D] | DBL: E F | DTE: X Y"
Dim B As FldTyMul: B = FldTyMulBrk(A)
Debug.Assert UBound(B.TyAy) = 2
Debug.Assert B.TyAy(0).Ty = DAO.dbText
Debug.Assert B.TyAy(1).Ty = DAO.dbDouble
Debug.Assert B.TyAy(2).Ty = DAO.dbDate
AyAsstEq B.TyAy(0).F, FnStrBrk("A B [C D]")
AyAsstEq B.TyAy(1).F, LvsSplit("E F")
AyAsstEq B.TyAy(2).F, LvsSplit("X Y")
End Sub

Function FldTyMulFny(A As FldTyMul) As String()
Dim U%: U = UBound(A.TyAy)
Dim J%, O$()
For J = 0 To U
    PushAy O, A.TyAy(J).F
Next
FldTyMulFny = O
End Function

Function FldTySngBrk(FldTySngStr) As FldTySng
Dim O As FldTySng
With StrBrk(FldTySngStr, ":")
    O.Ty = DaoTyNew(.S1)
    O.F = FnStrBrk(.S2)
End With
FldTySngBrk = O
End Function

Sub FldTySngBrk__Tst()
Dim A$, B As FldTySng
A = "DTE : A B C": B = FldTySngBrk(A): Debug.Assert B.Ty = dbDate:   AyAsstEq B.F, ApSy("A", "B", "C")
A = "DBL : A B C": B = FldTySngBrk(A): Debug.Assert B.Ty = dbDouble: AyAsstEq B.F, ApSy("A", "B", "C")
End Sub

Function FldTySngToStr$(A As FldTySng)
FldTySngToStr = FmtQQ("? : ?", DaoTyToStr(A.Ty), FnyToStr(A.F))
End Function
