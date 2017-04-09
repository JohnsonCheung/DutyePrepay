Attribute VB_Name = "nIde_nMth_MdLy"
Option Compare Database
Option Explicit

Function MdLyEIdxAy&(MdLy$(), MthTy$, FmI&)
Dim J&, A$
A = "End " & MthTy
For J = FmI To UB(MdLy)
    If IsPfx(MdLy(J), A) Then MdLyEIdxAy = J: Exit Function
Next
MdLyEIdxAy = -1
End Function

Function MdLyToMthBEIdxAy(MdLy$(), MthNm$, Optional PrpTy$, Optional FmI&) As Variant()
Dim DrAy()
Dim U&
    DrAy = MdLy_ToMthTy_BIdx_DrAy(MdLy, MthNm, PrpTy, FmI)
    U = UB(DrAy)

Dim O()
    ReSz O, U
    Dim J&, E&, Dr, MthTy$, B&
    For J = 0 To U
        Dr = DrAy(J)
        MthTy = Dr(1)
        B = Dr(0)
        E = MdLyEIdxAy(MdLy, MthTy, B + 1)
        O(J) = BEIdxNew(B, E)
    Next
MdLyToMthBEIdxAy = O
End Function

Function MdLyToMthBIdxAy(MdLy$(), MthNm$, Optional PrpTy$, Optional FmI&) As Long()
Dim M(): M = MdLy_ToMthTy_BIdx_DrAy(MdLy, MthNm, PrpTy, FmI)
MdLyToMthBIdxAy = DrAyCol_LngAy(M)
End Function

Private Function MdLy_ToMthTy_BIdx_DrAy(MdLy$(), MthNm$, Optional PrpTy$, Optional FmI&) As Variant()
Dim J&, M As MthBrk
Dim Sel As Boolean
Dim O(), L$
For J = 0 To UB(MdLy)
    If SrcLinIsMth(MdLy(J)) Then
        L = SrcLyOneContinueLin$(MdLy, J)
        M = MthBrkNew(L)
        If MthBrkMatch(M, MthNm, PrpTy) Then
            Push O, Array(J, M.Ty)
        End If
    End If
Next
MdLy_ToMthTy_BIdx_DrAy = O
End Function
