Attribute VB_Name = "nIde_nPrm_Md"
Option Compare Database
Option Explicit

Function MdMthPrm(Optional A As CodeModule) As MthPrm()
Dim Ay$(): Ay = MdMthLy(A)
Dim B As MthBrk
Dim MdN$: MdN = MdNm(A)
Dim M As Mth
Dim O() As MthPrm
Dim J%
For J = 0 To UB(Ay)
    B = MthBrkNew(Ay(J))
    Set M = Mth(MdN, B.Nm)
    PushObj O, MthPrm(M, PrmAy(B.PrmStr))
Next
MdMthPrm = O
End Function

Sub MdMthPrm__Tst()
Dim A() As MthPrm: A = MdMthPrm
AyBrw AyMapIntoSy(A, "MthPrmToStr")
End Sub
