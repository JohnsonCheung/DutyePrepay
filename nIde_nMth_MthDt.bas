Attribute VB_Name = "nIde_nMth_MthDt"
Option Compare Database
Option Explicit

Function MthDt_Md(Optional A As CodeModule) As Dt
Dim B$(): B = MdBdyLy(A)
Dim OD(), J%
For J = 0 To UB(B)
    If SrcLinIsMth(B(J)) Then
        Push OD, MthBrkDr(MthBrkNew(B(J)))
    End If
Next
MthDt_Md = DtNew(MthBrkFny, OD)
End Function

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

Sub MthDtBrw_Md(Optional A As CodeModule)
DtBrw MthDt_Md(A), MdNm(A)
End Sub

Sub MthDtBrw_Pj(Optional A As vbproject)
DtBrw MthDt_Pj(A), PjNm(A)
End Sub
