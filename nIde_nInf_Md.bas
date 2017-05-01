Attribute VB_Name = "nIde_nInf_Md"
Option Compare Database
Option Explicit

Function Md(Optional MdNm$, Optional Pj As vbproject) As CodeModule
If MdNm$ = "" Then
    Set Md = PjNz(Pj).Vbe.ActiveCodePane.CodeModule
    Exit Function
End If
Set Md = PjNz(Pj).VBComponents(MdNm).CodeModule
End Function

Function MdAllLines$(Optional A As CodeModule)
MdAllLines = MdLines(, , A)
End Function

Function MdAppa(Optional A As CodeModule) As Access.Application
'Find the Access.application of given {Md}.  Assume current
Set MdAppa = PjAppa(MdPj(A))
End Function

Function MdBdyLines$(Optional A As CodeModule)
With MdNz(A)
    Dim S&: S = .CountOfDeclarationLines + 1
    MdBdyLines = .Lines(S, .CountOfLines)
End With
End Function

Function MdBdyLinesSorted$(Optional A As CodeModule)
Dim AMth() As MthStru: AMth = MthStruAy_All(A)
Dim AIdx&():
    Dim Key$(): Key = MthStruAyKeyAy(AMth)
    AIdx = AySrtIdx(Key)
Dim ALy$(): ALy = MdLy(A)
Dim O$()
    Dim J&, Mth As MthStru
    Dim M$(), BIdx&, EIdx&
    For J = 0 To UB(AIdx)
        Mth = AMth(AIdx(J))
        BIdx = Mth.BEIdx(0)
        EIdx = Mth.BEIdx(1)
        M = AySlice(ALy, BIdx, EIdx)
        Push O, ""
        PushAy O, M
    Next
MdBdyLinesSorted = LyJn(O)
End Function

Function MdBdyLy(Optional A As CodeModule) As String()
Dim B As CodeModule: Set B = MdNz(A)
If B.CountOfLines = 0 Then Exit Function
Dim S&: S = B.CountOfDeclarationLines + 1
MdBdyLy = LinesSplit(B.Lines(S, B.CountOfLines))
End Function

Function MdCmp(Optional A As CodeModule) As VBComponent
Set MdCmp = MdNz(A).Parent
End Function

Function MdCur() As CodeModule
Dim A As CodePane: Set A = Application.Vbe.ActiveCodePane
If IsNothing(A) Then Exit Function
Set MdCur = A.CodeModule
End Function

Function MdCurNm$()
MdCurNm = MdNm(MdCur)
End Function

Function MdDclLines$(Optional A As CodeModule)
With MdNz(A)
    MdDclLines = .Lines(1, .CountOfDeclarationLines)
End With
End Function

Function MdDclLy(A As CodeModule) As String()
MdDclLy = LinesSplit(MdDclLines(A))
End Function

Function MdEnumNy(Optional A As CodeModule) As String()
Dim B$(): B = MdDclLy(A)
Dim O$(), J%
For J = 0 To UB(B)
    PushNoBlank O, SrcLinEnumNm(B(J))
Next
MdEnumNy = O
End Function

Function MdExt$(Optional A As CodeModule)
Dim B As CodeModule: Set B = MdNz(A)
Dim T As vbext_ComponentType: T = B.Parent.Type
Dim O$
Select Case T
Case vbext_ComponentType.vbext_ct_StdModule: O = ".bas"
Case vbext_ComponentType.vbext_ct_ClassModule: O = ".cls"
Case vbext_ComponentType.vbext_ct_MSForm: O = ".msfrm"
Case vbext_ComponentType.vbext_ct_Document: O = ".bas"
Case Else: Er "MdExt: {Ty} of given module is not in [STD | CLS | MSFORM | DOC]", T, MdNm(B)
End Select
MdExt = O
End Function

Function MdHasStr(Str$, Optional A As CodeModule) As Boolean
MdHasStr = InStr(MdAllLines(A), Str) > 0
End Function

Function MdIsInPj(MdNm$, Optional A As vbproject) As Boolean
MdIsInPj = AyHas(PjMdNy(A), MdNm)
End Function

Function MdIsSaved(Optional A As CodeModule) As Boolean
MdIsSaved = MdCmp(A).Saved
End Function

Function MdLines$(Optional BegLno& = 1, Optional Cnt& = 0, Optional A As CodeModule)
If BegLno <= 0 Then Exit Function
Dim Md As CodeModule: Set Md = MdNz(A)
Dim C&: C = Md.CountOfLines
If C = 0 Then Exit Function
Dim N&: N = IIf(Cnt > 0, Cnt, C)
MdLines = Md.Lines(BegLno, N)
End Function

Function MdLinesByBEIdx(BEIdx&(), Optional A As CodeModule)
Dim Fm&, N&: AyAsg BEIdxFmN(BEIdx, FmIsBase1:=True), Fm, N
MdLinesByBEIdx = MdNz(A).Lines(Fm, N)
End Function

Function MdLnoAftOptLin&(Optional A As CodeModule)
Dim Ly$(): Ly = MdDclLy(A)
Dim U&: U = UB(Ly)
If IsPfx(AyLasEle(Ly), "Option") Then MdLnoAftOptLin = U + 2: Exit Function
Dim J&
For J = 0 To UB(Ly)
    If Not IsPfx(Ly(J), "Option") Then MdLnoAftOptLin = J + 1: Exit Function
Next
MdLnoAftOptLin = 1
End Function

Function MdLy(Optional A As CodeModule, Optional BIdx& = -1, Optional EIdx& = -1) As String()
MdLy = LinesSplit(MdLines(, , A))
End Function

Function MdLyByBEIdx(BEIdx&(), Optional A As CodeModule) As String()
MdLyByBEIdx = LinesSplit(MdLinesByBEIdx(BEIdx, A))
End Function

Function MdMthLy(Optional A As CodeModule) As String()
MdMthLy = BdyLyMthLy(MdBdyLy(MdNz(A)))
End Function

Sub MdMthLy__Tst()
Pipe "Form_Switchboard", "MdMthLy", "AyBrw"
End Sub

Function MdNm$(Optional A As CodeModule)
MdNm = MdNz(A).Parent.Name
End Function

Function MdNz(A As CodeModule) As CodeModule
If IsNothing(A) Then
    Set MdNz = MdCur
Else
    Set MdNz = A
End If
End Function

Function MdOneLin$(BegLno&, Optional A As CodeModule)
MdOneLin = MdLines(BegLno, 1, A)
End Function

Function MdPgmDs(Optional A As CodeModule) As Ds
'
'Dim mNmtOldPgm$, mNmtOldArg$
'Do
'    Dim mA$(): mA = Split(Replace(pLnt, ":", CtComma), CtComma)
'    If Sz(mA) <> 2 Then ss.A 4, "pLnt must have 2 elements": GoTo E
'    mNmtOldPgm = Trim(mA(0))
'    mNmtOldArg = Trim(mA(1))
'    Dim mDPgm As New d_Pgm: If mDPgm.CrtTbl(mNmtOldPgm) Then ss.A 5: GoTo E
'    Dim mDArg As New d_Arg: If mDArg.CrtTbl(mNmtOldArg) Then ss.A 6: GoTo E
'    Dim mRsPgm As DAO.Recordset: Set mRsPgm = CurrentDb.TableDefs(mNmtOldPgm).OpenRecordset
'    Dim mRsArg As DAO.Recordset: Set mRsArg = CurrentDb.TableDefs(mNmtOldArg).OpenRecordset
'Loop Until True
'
'Dim iPrj%
'For iPrj = 0 To mNPrj - 1
'    Dim mNmPrj$: mNmPrj = mAnPrj(iPrj)
'    Dim mPrj As vbproject: If Fnd_Prj(mPrj, mNmPrj, mAcs) Then ss.A 8: GoTo E
'    Dim mAnm$(): If Fnd_Anm_ByPrj(mAnm, mPrj) Then ss.A 9: GoTo E
'
'    Dim iMd%
'    For iMd = 0 To Sz(mAnm) - 1
'        Dim mNmm$: mNmm = mAnm(iMd)
'        StsShw "Export pgm in module [" & mNmPrj & "." & mNmm & "]...."
'        Dim mMd As CodeModule: If Fnd_Md(mMd, mPrj, mNmm) Then ss.A 10: GoTo E
'        Dim mAnPrc$(): If Fnd_AnPrc_ByMd(mAnPrc, mMd, , True) Then ss.A 11: GoTo E
'        Dim iPrc%
'        For iPrc = 0 To Sz(mAnPrc) - 1
'            Dim mNmPrc$: mNmPrc = mAnPrc(iPrc):
'            If mNmPrc = "qBrkRec" And mNmm = "Brk" Then Stop
'            'Debug.Print "Prj(" & iPrj & ":" & mNmPrj & ") Md(" & iMd & ":" & mNmm & ") Prc(" & iPrc & ":" & mNmPrc & ")"
'            Dim mPrcBody$: If Fnd_PrcBody_ByMd(mPrcBody, mMd, mNmPrc, True) Then ss.A 12: GoTo E
'            Dim mAyDArg() As d_Arg
'            If Brk_PrcBody(mDPgm, mAyDArg, mPrcBody) Then ss.A 13: GoTo E
'            mDPgm.x_NmPrj = mNmPrj
'            mDPgm.x_Nmm = mNmm
'            If mDPgm.Ins(mRsPgm) Then ss.A 14: GoTo E
'            If mDArg.InsAy(mRsArg, mNmPrj, mNmm, mNmPrc, mAyDArg) Then ss.A 16: GoTo E
'        Next
'    Next
'Next
'GoTo X
'R: ss.R
'E: MdPgmDs = True: ss.B cSub, cMod, "pLnt,SrcFb", pLnt, SrcFb
'X:
'    If SrcFb <> "" Then Cls_CurDb mAcs
'    RsCls mRsPgm
'    RsCls mRsArg
'    Clr_Sts
End Function

Function MdPgmDs__Tst()
Const cSub$ = "MdPgmDs_Tst"
Dim mLnt$, mFbSrc$
Dim mCase As Byte: mCase = 1
Select Case mCase
Case 1
    mLnt$ = "#OldPgm,#OldArg"
    mFbSrc = "p:\workingdir\pgmobj\JMtcLgc.Mdb"
Case 2
    mLnt$ = "#OldPgm,#OldArg"
    mFbSrc = "p:\workingdir\pgmobj\JMtcDb.Mdb"
End Select
DsBrw MdPgmDs()
End Function

Function MdPj(A As CodeModule) As vbproject
Set MdPj = MdNz(A).Parent.Collection.Parent
End Function

Function MdSfx$(Optional A As CodeModule)
Dim B$: B = MdNm(A)
MdSfx = Brk2FmEnd(B, "_").S2
End Function

Function MdSrcFfn$(Optional A As CodeModule)
MdSrcFfn = PjSrcPth(MdPj(A)) & MdSrcFn(A)
End Function

Function MdSrcFn$(Optional A As CodeModule)
MdSrcFn = MdNm(A) & MdExt(A)
End Function

Function MdToStr$(A As CodeModule)
On Error GoTo R
MdToStr = PjNm(MdPj(A)) & "." & MdNm(A)
Exit Function
R: MdToStr = ErStr("MdToStr")
End Function

Function MdTy(Optional A As CodeModule) As vbext_ComponentType
MdTy = MdNz(A).Parent.Type
End Function

Function MdTyNy(Optional A As CodeModule) As String()
Dim B$(): B = MdDclLy(A)
Dim O$(), J%
For J = 0 To UB(B)
    PushNoBlank O, SrcLinTyNm(B(J))
Next
MdTyNy = O
End Function
