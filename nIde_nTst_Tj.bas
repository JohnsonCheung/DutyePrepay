Attribute VB_Name = "nIde_nTst_Tj"
Option Compare Database
Option Explicit

Function Tj(Optional Pj As vbproject) As vbproject
Dim P As vbproject: Set P = PjNz(Pj)
End Function

Sub TjCrt(Optional P As vbproject)
Dim P1 As vbproject: Set P1 = PjNz(P):            If PjIsTj(P1) Then Exit Sub
Dim pFil$:             pFil = P1.FileName
Dim TFil$:             TFil = TjFfn(P1)
PjCrtFfn TFil, ApSy(pFil)
End Sub

Function TjEns(Optional A As vbproject) As vbproject
'If Tj of module-A is open, just exist
'If Tj file not exit, create it
'If Tj is not open, open it
Dim B As vbproject: Set B = PjNz(A)
Dim C$:                 C = TjFfn(B)
                            If Not FfnIsExist(C) Then PjCrtFfn C
                Set TjEns = PjOpnFfn(C)
End Function

Sub TjEnsAllTm(Optional A As vbproject)
AyEachEle PjMdAy(PjNz(A)), "TmEns"
End Sub

Function TjFfn$(Optional P As vbproject)
Dim Pj As vbproject: Set Pj = PjNz(Pj)
Dim Pth$:               Pth = PjPth(Pj)
                        Pth = Pth & "Tst\": PthEns Pth
Dim F$:                   F = Pth & "Tst_" & Pj.Name & AppExt
TjFfn = F
End Function

Function TjIsOpn(Optional Pj As vbproject) As Boolean
Dim P As vbproject: Set P = PjNz(Pj)
Dim N$: N = "Tst_" & P.Name
TjIsOpn = PjIsOpn(N)
End Function

Sub TjLod(Optional Pj As vbproject)
Dim P As vbproject: Set P = PjNz(Pj): If PjIsTj(P) Then Exit Sub
Dim F$: F = TjFfn(P)
FfnAsstExist F, "TjLod {PjNm} {TjFfn}", P.Name, F
If PjIsOpn("Tst_" & P.Name) Then Exit Sub
Excel.Application.Workbooks.Open F
End Sub

Sub TjOpn(Optional Pj As vbproject)
Dim P As vbproject: Set P = PjNz(Pj)
If PjIsTj(P) Then Exit Sub
If TjIsOpn(P) Then Exit Sub
Dim F$: F = TjFfn(P)
Select Case OfcTy
Case eXls: Excel.Application.Workbooks.Open F
Case eAcs: Appa.OpenCurrentDatabase F
Case Else: Er "TjOpn: Invalid {OfcTy}", OfcTy
End Select
End Sub
