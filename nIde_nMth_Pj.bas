Attribute VB_Name = "nIde_nMth_Pj"
Option Compare Database
Option Explicit

Function PjMthNy(Optional A As vbproject) As String()
Dim MdAy() As CodeModule:
    MdAy = PjMdAy(A)
Dim Ay():
    Ay = AyMap(MdAy, "MdMthNy")
Dim O$()
Dim J&
For J = 0 To UB(MdAy)
    Dim Pfx$: Pfx = MdNm(MdAy(J)) & "."
    PushAy O, AyAddPfx(Ay(J), Pfx)
Next
PjMthNy = O
End Function

Sub PjMthNyBrw(Optional A As vbproject)
DrAyBrw AyBrk(PjMthNy(A)), BrkAtColIdx:=1
End Sub

Function PjMthNyNotMatchWithMdNm(Optional A As vbproject) As String()
'Each public method should have pfx matched with Sfx of module name
'If not, return these method names
Dim B As vbproject: Set B = PjNz(A)
Dim MdAy() As CodeModule: MdAy = PjMdAy(B)
Dim Ay$(), O$(), J%
For J = 0 To UB(MdAy)
    If MdTy(MdAy(J)) = vbext_ct_Document Then GoTo Nxt
    Ay = MdMthNyNotMatchWithMdNm(MdAy(J))
    Ay = AyAddPfx(Ay, MdNm(MdAy(J)) & ".")
    PushAy O, Ay
Nxt:
Next
PjMthNyNotMatchWithMdNm = O
End Function

Sub PjMthNyNotMatchWithMdNm__Tst()
DrAyBrw AyBrk(PjMthNyNotMatchWithMdNm, ".")
End Sub
