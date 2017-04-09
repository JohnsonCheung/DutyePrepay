Attribute VB_Name = "nIde_nMth_MthLin"
Option Compare Database
Option Explicit

Function MthLinEns$(MthLin$, A As CodeModule)
Dim oBrk As MthBrk:
    Dim MthNm$
    Dim NewMthNm$
    oBrk = MthBrkNew(MthLin)
    MthNm = oBrk.Nm
    NewMthNm = MthNmEns(MthNm, A)          ' Bm = Module Methods    ! given module all methods
    oBrk.Nm = NewMthNm
MthLinEns = MthBrkToStr(oBrk)
End Function

Function MthLinkTakPrm__Tst()
Debug.Assert MthLinTakPrm("lksdjf()lksdjf,(lskdjf, dg 1)klsdj") = "lskdjf, dg 1"
End Function

Function MthLinRen$(MthLin, NewNm$)
Dim A As MthBrk: A = MthBrkNew(MthLin)
A.Nm = NewNm
MthLinRen = MthBrkToStr(A)
End Function

Function MthLinTakPrm$(MthLin$)
'Aim: Cut the {oPrm} out from S
Dim mA$: mA = Replace(MthLin, "()", "  ")
Dim mP1%: mP1 = InStr(mA, "(")
Dim mP2%: mP2 = InStr(mA, ")")
If mP1 = 0 Then Exit Function
If mP2 = 0 Then Exit Function
If mP1 > mP2 Then Exit Function
MthLinTakPrm = Mid(MthLin, mP1 + 1, mP2 - mP1 - 1)
End Function
