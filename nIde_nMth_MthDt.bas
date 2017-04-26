Attribute VB_Name = "nIde_nMth_MthDt"
Option Compare Database
Option Explicit

Function MthDt_Md(Optional A As CodeModule) As Dt
Dim B$(): B = MdMthLy(A)
Dim OD()
Dim J%
For J = 0 To UB(B)
    Push OD, MthBrkDr(MthBrkNew(B(J)))
Next
MthDt_Md = DtNew(MthBrkFny, OD)
End Function

Sub MthDt_Md__Tst()
DtBrw MthDt_Md
End Sub

Function MthDt_Pj(Optional A As vbproject) As Dt
Dim O()
Dim MdAy() As CodeModule: MdAy = PjMdAy(A)
If AyIsEmpty(MdAy) Then Exit Function

Dim I, Md As CodeModule, M As Dt
For Each I In MdAy
    Set Md = I
    M = MthDt_Md(Md)
    PushAy O, DrAyAddCol_Const(M.DrAy, MdNm(Md))
Next
Dim OF$(): OF = M.Fny: Push OF, "MdNm"
MthDt_Pj = DtNew(OF, O)
End Function

Sub MthDt_Pj__Tst()
DtBrw MthDt_Pj
End Sub

Sub MthDtBrw_Md(Optional A As CodeModule)
DtBrw MthDt_Md(A), MdNm(A)
End Sub

Sub MthDtBrw_Pj(Optional A As vbproject)
DtBrw MthDt_Pj(A), PjNm(A)
End Sub
