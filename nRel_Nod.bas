Attribute VB_Name = "nRel_Nod"
Option Compare Database
Option Explicit

Sub NodAsstVdt(Nod())
ErAsst NodChkVdt(Nod)
End Sub

Function NodChkVdt(Nod()) As Variant()
Dim O()
If Sz(Nod) <> 2 Then O = ErNew("Given Nod-{Sz} should be 2", Sz(Nod)): Exit Function
If Not VarIsStr(Nod(0)) Then O = ErNew("Given Nod(0)-{Ty} should be Str", TypeName(Nod(0)))
If Not VarIsSy(Nod(1)) Then O = ErApd(O, "Given Nod(1)-{Ty} should be StrAy", TypeName(Nod(1)))
If Trim(Nod(0)) = "" Then O = ErApd(O, "Given Nod(0) should not be *Blank")
If AyHasDup(Nod(1)) Then O = ErApd(O, "Given Nod(1) should not have dup item", Jn(AyDupItm(Nod(1)), " "))
NodChkVdt = O
End Function

Function NodIsEr(Nod()) As Boolean
NodIsEr = True
If Sz(Nod) <> 2 Then Exit Function
If Not VarIsStr(Nod(0)) Then Exit Function
If Not VarIsSy(Nod(1)) Then Exit Function
If Trim(Nod(0)) = "" Then Exit Function
If AyHasDup(Nod(1)) Then Exit Function
NodIsEr = False
End Function

Function NodNew(RelLin) As Variant()
Dim O(1)
With Brk(RelLin, ":")
    O(0) = .S1
    O(1) = LvsSplit(.S2)
End With
NodAsstVdt O
NodNew = O
End Function

Sub NodNew__Tst()
Dim A()
A = NodNew("A : B C D")
Dim Par$, Chd$()
Par = A(0)
Chd = A(1)
Debug.Assert Par = "A"
Debug.Assert Sz(Chd) = 3
Debug.Assert Chd(0) = "B"
Debug.Assert Chd(1) = "C"
Debug.Assert Chd(2) = "D"
End Sub

Function NodToStr$(Nod())
NodToStr = Nod(0) & " : " & Jn(Nod(1), " ")
End Function
