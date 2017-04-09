Attribute VB_Name = "nIde_nTst_Tm"
Option Compare Database
Option Explicit

Sub TmBld(Optional A As CodeModule)
If NmIsTstNm(MdNm(A)) Then Exit Sub
End Sub

Function TmEns(Optional A As CodeModule) As CodeModule
If MdIsTstNm(A) Then Er "Given {Md} has Pfx [Tst_] cannot be used to TmEns", MdNm(A)
Dim OMd As CodeModule: Set OMd = MdNz(A)
Dim OPj As vbproject:  Set OPj = TjEns(MdPj(OMd))
Dim OMdNm$
    OMdNm = NmToTstNm(MdNm(Md))
If Not MdIsInPj(OMdNm, OPj) Then MdCrt OMdNm, Pj:=OPj
Set TmEns = Md(OMdNm, OPj)
End Function
