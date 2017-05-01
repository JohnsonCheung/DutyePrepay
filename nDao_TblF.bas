Attribute VB_Name = "nDao_TblF"
Option Compare Database
Option Explicit

Sub TblFldAsstFldTyMulStr(T$, FldTyAllStr$, Optional A As database)
ErAsst TblFldChkFldTyMulStr(T, FldTyAllStr, A)
End Sub

Sub TblFldAsstFldTyMulStr__Tst()
Const T$ = "#Tmp"
Const F$ = "TXT: AA | INT : BB"
TblDrp T
TblCrt T, "AA Text, BB Integer"
TblFldAsstFldTyMulStr T, F
TblDrp T
End Sub

Sub TblFldAsstFldTySngStr(T$, FldTySngStr$, Optional A As database)
ErAsst TblFldChkFldSngStr(T, FldTySngStr, A)
End Sub

Function TblFldChkFldSngStr(T$, FldTySngStr$, Optional A As database) As Variant()
Dim D As database:         Set D = DbNz(A)
Dim A1 As FldTySng:        A1 = FldTySngBrk(FldTySngStr)
Dim Ty As DAO.DataTypeEnum:     Ty = A1.Ty
Dim B1$():                    B1 = TblFny(T, , D)
Dim B$():                      B = AyIntersect(B1, A1.F)
Dim C1() As DAO.Field:        C1 = TblFldAy(T, B, D)
Dim C() As DAO.DataTypeEnum:   C = OyPrp_Into(C1, "TYPE", C)

Dim O$(), O1() As DataTypeEnum
    Dim J%
    For J = 0 To UB(B)
        If C(J) <> Ty Then
            Push O, B(J)
            Push O1, C(J)
        End If
    Next

If AyIsEmpty(O1) Then Exit Function
Dim OEr(): OEr = ErNew("Following fields of {Table} should have this {DtaTy}", T, DaoTyToStr(Ty))
For J = 0 To UB(O)
    OEr = ErApd(OEr, "." & J, O(J), DaoTyToStr(O1(J)))
Next
TblFldChkFldSngStr = OEr
End Function

Function TblFldChkFldTyMulStr(T$, FldTyMulStr$, Optional A As database) As Variant()
Dim F As FldTyMul:   F = FldTyMulBrk(FldTyMulStr)
Dim Fny$():        Fny = FldTyMulFny(F)
Dim Fs$:            Fs = FnyToStr(Fny)
Dim O():             O = TblFldChkFnStr(T, Fs, A)
Dim J%
For J = 0 To UBound(F.TyAy)
    O = AyAdd(O, TblFldChkFldSngStr(T, FldTySngToStr(F.TyAy(J)), A))
Next
TblFldChkFldTyMulStr = O
End Function

Sub TblFldChkFldTyMulStr__Tst()
Const T$ = "#Tmp"
TblDrp T
TblCrt T, "AA TEXT(10),BB SHORT"
Dim Act(): Act = TblFldChkFldTyMulStr(T, "TXT : AA | INT : BB")
Debug.Assert Not AyHasEle(Act)
TblDrp T
End Sub

Function TblFldChkFnStr(T$, FnStr$, Optional D As database) As Variant()
Dim F$(): F = NmBrk(FnStr)
TblFldChkFnStr = TblFldChkFny(T, F, D)
End Function

Function TblFldChkFny(T$, Fny$(), Optional D As database) As Variant()
Dim A$(): A = Fny
Dim B$(): B = TblFny(T, , D)
Dim C$(): C = AyMinus(A, B)
If AyIsEmpty(C) Then Exit Function
TblFldChkFny = ErNew("{Tbl} with these {fields} has {missing-fields}", T, FnyToStr(B), FnyToStr(C))
End Function
