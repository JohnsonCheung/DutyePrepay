Attribute VB_Name = "nIde_nPj_nDo_Pj"
Option Compare Database
Option Explicit

Function PjAddCmp(CmpNm$, Optional CmpTy As vbext_ComponentType = vbext_ct_StdModule, Optional A As vbproject) As VBComponent
Dim O As VBComponent
Set O = PjNz(A).VBComponents.Add(CmpTy)
O.Name = CmpNm
End Function

Function PjAddCmp__Tst()
Const CmpNm$ = "AAAA"
Dim Act As VBComponent:
Set Act = PjAddCmp(CmpNm)
PjRmvCmp CmpNm
End Function

Sub PjAddRf(RfFfn, A As vbproject)
PjNz(A).References.AddFromFile RfFfn
End Sub

Sub PjAddRf__Tst()
Dim mCase As Byte
mCase = 2
Dim mFfn$, mPrj As vbproject:
Dim mAcs As Access.Application: Set mAcs = G.gAcs
Dim mFb$: mFb = "c:\tmp\aa.mdb": FbNew mFb
If Opn_CurDb(mAcs, mFb) Then Stop:

Select Case mCase
Case 1
    Set mPrj = mAcs.Vbe.ActiveVBProject
    Dim mRf As VBIDE.Reference: Set mRf = mPrj.References.AddFromFile("c:\program files\sap\frontend\sapgui\awkone.ocx")

    Dim mNmPrj$: mNmPrj = mPrj.Name
    mFfn = "C:\tmp\aa.reference.txt"
    PjRfDt mFfn, mNmPrj, mAcs

    mPrj.References.Remove mRf
Case 2
    mFfn = "P:\Documents\Pgm\Reference.txt"
    Set mPrj = mAcs.Vbe.ActiveVBProject
End Select
Stop
If Add_Rf(mPrj, mFfn) Then Stop
mAcs.Visible = True
Stop
End Sub

Sub PjBrwSrcPth(Optional A As vbproject)
PthBrw PjSrcPth(PjNz(A))
End Sub

Function PjCmp(CmpNm$, Optional A As vbproject) As VBComponent
Set PjCmp = PjNz(A).VBComponents(CmpNm)
End Function

Sub PjCpyMd(FmPj As vbproject, Optional MdNmLik$ = "*", Optional ToMdNmPfx = "ZZ_", Optional ToPj As vbproject)
Dim MdAy() As CodeModule: MdAy = PjMdAy(FmPj, MdNmLik)
AyEachEle MdAy, "MdCpy", PjNz(ToPj), ToMdNmPfx
End Sub

Function PjCrtFfn(PjFfn$, Optional RfFfnAy) As vbproject
Dim O As vbproject
FfnAsstNotExist PjFfn, "Cannot PjCrtFfn"
Dim Ext$: Ext = FfnExt(PjFfn)
Select Case Ext
Case AppxExt: Set O = FxlamCrt(PjFfn)
'Case AppwExt: Set O = FmdaCrt(PjFfn)
'Case AppoExt: Set O = CrtPowFfn(PjFfn)
'Case ApppExt: Set O = CrtPjoFfn(PjFfn)
Case AppaExt: Set O = FmdaCrt(PjFfn)
Case Else: Er "PjCrtFfn: Invalid {Ext} of {PjFfn}", Ext, PjFfn
End Select
Set PjCrtFfn = O
Stop
'AddPjRfFfnAy O, SyOpt(RfFfnAy)
End Function

Sub PjCrtFfn__Tst()
PjCrtFfn "c:\temp\a.xlam"
End Sub

Sub PjExp(Optional A As vbproject)
Dim P As vbproject: Set P = PjNz(A)
PthClr PjSrcPth(P)
AyEachEle PjMdAy(P), "MdExp"
End Sub

Function PjOpnFfn(PjFfn$) As vbproject
Dim O As vbproject
Dim Ext$: Ext = FfnExt(PjFfn)
Select Case Ext
Case AppxExt: Set O = OpnPjFxlam(PjFfn)
Case AppaExt: Set O = AppaOpnPjFmda(PjFfn)
'Case AppExtOfWrd: Set O = OpnPjOfWrd(PjFfn)
'Case AppExtOfPpt: Set O = OpnPjOfPt(PjFfn)
'Case AppExtOfOlk:: Set O = OPnPjOfOlk(PjFfn)
Case Else: Er "Invald {Ext} of {PjFfn}, cannot OpnPjFfn", Ext, PjFfn
End Select
Set PjOpnFfn = O
End Function

Sub PjRmk(A As vbproject)
Dim MdAy() As CodeModule: MdAy = PjMdAy(A)
AyEachEle MdAy, "MdRmk"
End Sub

Sub PjRmvCmp(CmpNm$, Optional CmpTy As vbext_ComponentType = vbext_ct_StdModule, Optional A As vbproject)
PjNz(A).VBComponents.Remove PjCmp(CmpNm, A)
End Sub

Sub PjRmvEmptyMd(Optional A As vbproject)
'Dim N$(): N = PjEmptyMdNy
End Sub

Sub PjSav(Optional A As vbproject)
Dim C As VBComponent
For Each C In PjNz(A).VBComponents
    If Not C.Saved Then
        MdSav C.CodeModule
    End If
Next
End Sub

Sub PjSrt(Optional A As vbproject)
AyEachEle PjMdAy(A), "MdSrt"
End Sub
