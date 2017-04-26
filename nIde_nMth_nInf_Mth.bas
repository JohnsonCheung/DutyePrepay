Attribute VB_Name = "nIde_nMth_nInf_Mth"
Option Compare Database
Option Explicit

Sub MthAsstInMd(MthNm$, Optional A As CodeModule, Optional ErDr)
If Not MthIsInMd(MthNm, A) Then Er "MthAssInMd: Given {MthNm} not in {Md}", MthNm, MdNm(A)
End Sub

Function MthBEIdxAy(MthNm$, Optional PrpTy$, Optional A As CodeModule) As Variant()
Dim Md As CodeModule: Set Md = MdNz(A)
Dim Ly$(): Ly = MdLy(Md)
MthBEIdxAy = MdLyToMthBEIdxAy(Ly, MthNm, PrpTy, Md.CountOfDeclarationLines)
End Function

Function MthBIdxAy(Optional MthNm$, Optional PrpTy$, Optional A As CodeModule) As Long()
Dim Md As CodeModule: Set Md = MdNz(A)
Dim Ly$(): Ly = MdLy(Md)
MthBIdxAy = MdLyToMthBIdxAy(Ly, MthNm, PrpTy, Md.CountOfDeclarationLines)
End Function

Function MthIsInMd(MthNm$, Optional A As CodeModule) As Boolean
MthIsInMd = AyHas(MdMthNy(A), MthNm)
End Function

Function MthLin$(MthNm$, Optional A As CodeModule)
Dim Md As CodeModule: Set Md = MdNz(A)
Dim L&: L = MthLno(MthNm, Md)
If L = 0 Then Exit Function
MthLin = Md.Lines(L, 1)
End Function

Function MthLines$(MthNm$, Optional PrpTy$, Optional A As CodeModule)
Dim Md As CodeModule: Set Md = MdNz(A)
Dim Stru() As MthStru: Stru = MthStruAy(MthNm, PrpTy, A)
If MthStruAyIsEmpty(Stru) Then Exit Function
Dim J%, O$(), M As MthStru
For J = 0 To UBound(Stru)
    M = Stru(J)
    Push O, MdLinesByBEIdx(M.BEIdx, Md)
Next
MthLines = LyJn(O)
End Function

Function MthLno%(Optional MthNm$, Optional A As CodeModule)
Dim N$: N = MthNmNz(MthNm, A)
Dim Ly$(): Ly = MdBdyLy(A)
Dim J&
For J = 0 To UB(Ly)
    If SrcLinIsMth(Ly(J)) Then
        MthLno = J + Md.CountOfDeclarationLines + 1
        Exit Function
    End If
Next
End Function

Sub MthLno__Tst()
Debug.Assert MthLno("MthLno", Md("nIde_nMth_nInf_Mth")) = 123
End Sub

Function MthLy(MthNm$, Optional PrpTy$, Optional A As CodeModule) As String()
MthLy = LinesSplit(MthLines(MthNm, PrpTy, A))
End Function
