Attribute VB_Name = "nIde_nMth_Md"
Option Compare Database
Option Explicit

Function MdHasMth(A As CodeModule, MthNm$) As Boolean
Dim OMthNy$(): OMthNy = MdMthNy(A)
MdHasMth = AyHas(OMthNy, MthNm)
End Function

Function MdHasMth_Pub(A As CodeModule, PubMthNm$) As Boolean
Dim OMthNy$(): OMthNy = MdMthNyPub(A)
MdHasMth_Pub = AyHas(OMthNy, PubMthNm)
End Function

Function MdHasNoMth(Optional A As CodeModule) As Boolean
Dim B$(): B = MdLy(A)
Dim J%
For J = 0 To UB(B)
    If SrcLinIsMth(B(J)) Then Exit Function
Next
MdHasNoMth = True
End Function

Function MdIsEmpty(Optional A As CodeModule) As Boolean
MdIsEmpty = LinesTrim(MdBdyLines(A)) = ""
End Function

Function MdMthNy(Optional A As CodeModule) As String()
MdMthNy = MdMthNyMfy("", A)
End Function

Function MdMthNyMfy(mFY$, Optional A As CodeModule) As String()
Dim Md As CodeModule: Set Md = MdNz(A)
Dim B$(): B = MdBdyLy(Md)
Dim O$(), J%, L$, Brk As MthBrk, IMfy$
For J = 0 To UB(B)
    If SrcLinIsMth(B(J)) Then
        L = MdOneContinueLin(J + 1 + Md.CountOfDeclarationLines, Md)
        Brk = MthBrkNew(L)
        If mFY = "" Then
            Push O, Brk.Nm
        Else
            IMfy = Brk.mFY
            If mFY = "Public" Then
                If IMfy = "" Or IMfy = "Public" Then Push O, Brk.Nm
            Else
                If IMfy = mFY Then
                    Push O, Brk.Nm
                End If
            End If
        End If
    End If
Next
MdMthNyMfy = O
End Function

Function MdMthNyNotMatchWithMdNm(Optional A As CodeModule) As String()
'Each public method should have pfx matched with Sfx of module name
'If not, return these method names
Dim B As CodeModule: Set B = MdNz(A)
Dim mSfx$: mSfx = MdSfx(B)
Dim MthNy$(): MthNy = MdMthNyPub(B)
MdMthNyNotMatchWithMdNm = AyExcl(MthNy, "IsPfx", mSfx)
End Function

Function MdMthNyPri(Optional A As CodeModule) As String()
MdMthNyPri = MdMthNyMfy("Private", A)
End Function

Function MdMthNyPub(Optional A As CodeModule) As String()
MdMthNyPub = MdMthNyMfy("Public", A)
End Function

Function MdOneContinueLin$(Lno&, _
    Optional A As CodeModule)
Dim Md As CodeModule: Set Md = MdNz(A)
Dim O$(), J&, L$
    For J = Lno To Lno + 100
        L = Md.Lines(J, 1)
        If LasChr(L) <> "_" Then
            Push O, L
            MdOneContinueLin = Join(O)
            Exit Function
        End If
        Push O, Trim(RmvLasChr(L))
    Next
    Er "MdOneContinueLin: Impossible"
End Function

Sub MdOneContinueLin__Tst()
Const MthNm$ = "MdOneContinueLin"
Debug.Assert MdOneContinueLin(MthLno(MthNm), Md("nIde_nMth_Md")) = "Function MdOneContinueLin(Lno&, Optional A CodeModule)"
End Sub

