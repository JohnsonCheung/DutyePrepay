Attribute VB_Name = "nIde_nSrc_DimLin"
Option Compare Database
Option Explicit

Function DimLinNy(DimLin) As String()
Dim L$: L = DimLinTrim(DimLin)
If Not IsPfx(L, "Dim ") Then Er "Given {DimLin} is not Pfx-[Dim ]", DimLin
DimLinNy = PrmStrToNy(RmvPfx(L, "Dim "))
End Function

Sub DimLinNy__Tst()
Dim A$(), B$(), DimLy$()
Dim I
Dim Md As CodeModule
For Each I In PjMdAy
    Set Md = I
    A = MdLy(Md)
    B = SrcLyDimLy(A)
    PushAy DimLy, B
Next
Dim DrAy()
For Each I In DimLy
    A = DimLinNy(I)
    Push DrAy, ApSy(Trim(I), A)
Next
DrAyBrw DrAy
End Sub

Function DimLinTrim$(DimLin)
DimLinTrim = Brk1(SrcLinRmvRmk(Trim(DimLin)), ":").S1
End Function
