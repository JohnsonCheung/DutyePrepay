Attribute VB_Name = "nIde_nDo_Md"
Option Compare Database
Option Explicit

Sub MdApdLines(Lines, Optional A As CodeModule)
With MdNz(A)
    .InsertLines .CountOfLines + 1, Lines
End With
End Sub

Sub MdBdyBrw(Optional A As CodeModule)
StrBrw MdBdyLines(A)
End Sub

Function MdBEIdx(Optional A As CodeModule) As Long()
MdBEIdx = BEIdxNew(0, MdNz(A).CountOfLines - 1)
End Function

Sub MdBrw(Optional A As CodeModule)
StrBrw MdAllLines(A), MdNm(A)
End Sub

Sub MdBrwNm(Nm$, Optional Pj As vbproject)
MdBrw Md(Nm, Pj)
End Sub

Sub MdClr(Optional A As CodeModule)
With MdNz(A)
    If .CountOfLines > 0 Then
        .DeleteLines 1, .CountOfLines
    End If
End With
End Sub

Sub MdCls(Optional A As CodeModule)
Dim Md As CodeModule: Set Md = MdNz(A)
If IsNothing(Md.CodePane.Window) Then Exit Sub
Md.CodePane.Window.Close
End Sub

Sub MdCpy(FmMd As CodeModule, Optional ToPj As vbproject, Optional NewMdNmPfx$)
Dim ONewNm$
Dim OLines$
Dim OTy As vbext_ComponentType
Dim OToPj As vbproject
    ONewNm = NewMdNmPfx & MdNm(FmMd)
    OLines = MdAllLines(FmMd)
    OTy = MdTy(FmMd)
    Set OToPj = PjNz(ToPj)
MdCrt ONewNm, OLines, OTy, OToPj
End Sub

Sub MdCrt(MdNm$, Optional Lines$, Optional Ty As vbext_ComponentType = vbext_ct_StdModule, Optional Pj As vbproject)
Dim Md As CodeModule
    Set Md = PjNz(Pj).VBComponents.Add(Ty).CodeModule
Md.Name = MdNm
If Lines = "" Then Exit Sub
MdClr Md
MdApdLines Lines, Md
End Sub

Sub MdDltLin(BEIdx&(), Optional A As CodeModule)
Dim B&, E&: BEIdxAsg BEIdx, B, E
MdNz(A).DeleteLines B + 1, E - B + 1
End Sub

Sub MdExp(A As CodeModule)
A.Parent.Export MdSrcFfn(A)
End Sub

Sub MdOpn(A As CodeModule)
A.Parent.Activate
End Sub

Sub MdOpnNm(Nm$, Optional Pj As vbproject)
MdOpn Md(Nm, Pj)
End Sub

Sub MdRen(Fm$, ToNm$, Optional Pj As vbproject)
Dim P As vbproject: Set P = PjNz(Pj)
If MdIsInPj(ToNm, P) Then Er "{ToMdNm} exist in {P}", ToNm, PjNm(P)
P.VBComponents(Fm).Name = ToNm
End Sub

Sub MdRen__Tst()
MdRenPfx "nOfc_", ""
End Sub

Sub MdRenPfx(FmPfx$, ToPfx$, Optional Pj As vbproject)
Dim P As vbproject: Set P = PjNz(Pj)
Dim FmAy1$(): FmAy1 = PjMdNy(P)                         ' FmAy = From Module-Name-Array
Dim FmAy$():   FmAy = AyLik(FmAy1, FmPfx & "*")
Dim ToNm$, J%, ToNm1$
For J = 0 To UB(FmAy)
    ToNm1 = ToPfx & RmvPfx(FmAy(J), FmPfx)
    ToNm = PjNxtMdNm(ToNm1, P)
    MdRen FmAy(J), ToNm, P
Next
End Sub

Sub MdRmk(Optional A As CodeModule)
Dim Md As CodeModule: Set Md = MdNz(A)
MdRpl LyJn(AyAddPfx(MdLy(Md), "'")), Md
End Sub

Sub MdRmv(A As CodeModule)
Dim Pj As vbproject
Set Pj = MdPj(A)
Pj.VBComponents.Remove MdCmp(A)
End Sub

Sub MdRmvBdy(A As CodeModule)
Dim B As CodeModule: Set B = MdNz(A)
Dim S&: S = B.CountOfDeclarationLines + 1
Dim C&: C = B.CountOfLines - B.CountOfDeclarationLines
B.DeleteLines S, C
End Sub

Sub MdRmvLines(BEIdx&(), Optional A As CodeModule)
Dim Fm&, N&: AyAsg BEIdxFmN(BEIdx, FmIsBase1:=True), Fm, N
MdNz(A).DeleteLines Fm, N
End Sub

Sub MdRmvOneLin(Lno&, Optional A As CodeModule)
MdNz(A).DeleteLines Lno
End Sub

Sub MdRpl(MdLines$, Optional A As CodeModule)
Dim Md As CodeModule
    Set Md = MdNz(A)
MdRmvLines MdBEIdx(Md), Md
MdApdLines MdLines, A
End Sub

Sub MdRplBdy(Bdy$, Optional A As CodeModule)
MdRmvBdy A
MdApdLines Bdy, A
End Sub

Sub MdRplLin(Idx&, Lin$, Optional A As CodeModule)
With MdNz(A)
    .ReplaceLine Idx + 1, Lin
End With
End Sub

Sub MdRplStr(Fm$, ToStr$, Optional A As CodeModule)
Dim B As CodeModule: Set B = MdNz(A)
Dim Bdy$: Bdy = MdBdyLines(B)
If InStr(Bdy, Fm) = 0 Then Exit Sub
Dim O$: O = Replace(Bdy, Fm, ToStr)
MdRplBdy O, B
End Sub

Sub MdRplStr__Tst()
MdRplStr "ABC" & "XYZ", "123456", Md("nIde_nTth_Md")
End Sub

Sub MdSav(Optional A As CodeModule)
If MdIsSaved(A) Then Exit Sub
Dim Nm$: Nm = MdNm(A)
Dim Ty As Access.AcObjectType
    Select Case MdTy(A)
    Case vbext_ct_StdModule, vbext_ct_ClassModule
        Ty = Access.AcObjectType.acModule
    Case vbext_ct_Document
        Dim B$: B = TakBef(Nm, "_")
        Select Case B
        Case "Form": Ty = acForm
        Case "Report": Ty = acReport
        Case Else: Er "Unexpect {MdNm} {Pfx} which has type-of-Document", MdNm(A), B
        End Select
    Case Else: Er "Unexpect {MdTy} of {MdNm}", MdTy(A), MdNm(A)
    End Select
Dim App As Access.Application: Set App = MdAppa(A)
App.DoCmd.Save Ty, Nm
Debug.Print Nm & " <== Saved"
End Sub

Sub MdSavNm(Nm$)
MdSav Md(Nm)
End Sub

Sub MdSelLCC(P As LCC, Optional A As CodeModule)
With P
    MdNz(A).CodePane.SetSelection .L, .C1, .L, .C2
End With
End Sub

Sub MdSelTxt(P As LCLC, Optional A As CodeModule)
With P
    MdNz(A).CodePane.SetSelection .L1, .C1, .L2, .C2
End With
End Sub

Sub MdShw(A As CodeModule)
A.CodePane.Show
End Sub

Sub MdShwLno(Lno&, Optional A As CodeModule)
Dim Md  As CodeModule
Dim W As VBIDE.CodePane
Set Md = MdNz(A)
With Md.CodePane
    .SetSelection Lno, 1, Lno, 1
    .Show
    .Window.SetFocus
End With
End Sub

Sub MdSrt(Optional A As CodeModule)
Dim AMd As CodeModule: Set AMd = MdNz(A)
Dim ANm$: ANm = MdNm(AMd)

Debug.Print AlignL(ANm, 30);
If ANm = "nIde_nMd_nDo_Md" Then GoTo X
If ANm = "nMGI_UsrPrf" Then GoTo X

Dim AIsNoMth As Boolean: AIsNoMth = MdHasNoMth(AMd)
If MdHasNoMth(AMd) Then Debug.Print "<== no methond": Exit Sub

Dim ANewBdy$:
Dim AOldBdy$
    ANewBdy = LinesTrimEnd(MdBdyLinesSorted(AMd))
    AOldBdy = LinesTrimEnd(MdBdyLines(AMd))

If ANewBdy = AOldBdy Then
    Debug.Print "<== no change after sorted"
    Exit Sub
End If
MdRplBdy ANewBdy, AMd
Debug.Print "*** SORTED ***"
'MdSav AMd
Exit Sub
X: Debug.Print "<-- Skip"
End Sub

Sub MdSrtNm(MdNm$)
MdSrt Md(MdNm)
End Sub
