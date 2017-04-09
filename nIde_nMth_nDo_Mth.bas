Attribute VB_Name = "nIde_nMth_nDo_Mth"
Option Compare Database
Option Explicit

Sub MthApd(MthLines$, Optional A As CodeModule)
Dim OMd As CodeModule
    Set OMd = MdNz(A)

Dim Exist As Boolean:
    Exist = MdHasStr(MthLines, OMd)
If Exist Then Exit Sub

Dim OMthLy$()
    Dim MthLin$
    MthLin = OMthLy(0)
    OMthLy = LinesSplit(MthLines)
    OMthLy(0) = MthLinEns(MthLin, OMd)
    
MdApdLines LyJn(OMthLy), OMd
End Sub

Sub MthRen(Fm$, ToMthNm$, Optional A As CodeModule)
Dim OMd As CodeModule
    Set OMd = MdNz(A)
'-----
Dim BNewMthNm$
Dim BLy$()
    BNewMthNm = MthNmEns(ToMthNm, A)
    BLy = MdLy(OMd)
'-----
Dim ANewMthLy$()
Dim ABIdxAy&()
    Dim MthLinAy$()
    ABIdxAy = MdLyToMthBIdxAy(BLy, Fm)
    MthLinAy = AySel_Idx(BLy, ABIdxAy)
    ANewMthLy = AyMapIntoSy(MthLinAy, "MthLinRen", ToMthNm)
'-----
Dim OIdxLinDrAy()
    OIdxLinDrAy = DrAyNew_AyAp(ABIdxAy, ANewMthLy)
'----
Dim J%, Idx&, Lin$, Dr
For J = 0 To UB(OIdxLinDrAy)
    Dr = OIdxLinDrAy(J)
    AyAsg Dr, Idx, Lin
    MdRplLin Idx, Lin, OMd    '<==
Next
End Sub

Sub MthRen__Tst()
MthRen "AAA__Tst", "AA1"
End Sub

Sub MthRmv(MthNm$, Optional A As CodeModule)
Dim Md As CodeModule: Set Md = MdNz(A)
Dim B() As MthStru:
B = MthStruAy(MthNm, , Md)
If MthStruAyIsEmpty(B) Then Exit Sub
Dim J%
For J = UBound(B) To 0 Step -1
    MdRmvLines B(J).BEIdx, Md
Next
End Sub

Sub MthRmv__Tst()
MthApd RplVBar("Sub XXX()|End Sub")
End Sub

Sub MthShw(MthNm$, Optional A As CodeModule)
Dim Md As CodeModule: Set Md = MdNz(A)
MthAsstInMd MthNm, Md, "MthShw"
Dim P As LCC
With MdNz(A).CodePane
    .Show
    MdSelLCC MthLCC(MthNm, Md), Md
End With
End Sub

Sub MthShw__Tst()
MthShw "AAA__Tst"
End Sub

Sub MthShw_ForEdt(MthNm$, Optional A As CodeModule)
MthAsstInMd MthNm, A, "MthShw"
Dim P As LCC
With MdNz(A).CodePane
    .Show
    MdSelLCC MthLCC_ForEdt(MthNm, A), A
End With
End Sub

Sub MthStruAyKeyAy__Tst()
Dim A$(): A = MthStruAyKeyAy(MthStruAy(Md("mMd")))
AyBrw A
End Sub
