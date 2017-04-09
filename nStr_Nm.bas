Attribute VB_Name = "nStr_Nm"
Option Compare Database
Option Explicit

Function NmIsTstNm(Nm$) As Boolean
NmIsTstNm = StrHas(Nm, "Tst_") Or StrHas(Nm, "_Tst")
End Function

Function NmNxt$(Nm$, Ny$())
If Not AyHas(Ny, Nm) Then NmNxt = Nm: Exit Function
Dim A$: A = Nm & "_*"
Dim B$(): B = AySel(Ny, "Lik", A)

Dim J%, C$
For J = 0 To 100
    C = Nm & "_" & J
    If Not AyHas(B, C) Then NmNxt = C: Exit Function
Next
Er "NmNxt: Impossible"
End Function

Function NmTo2DashSfxTst$(TthNm$)
Dim A$: A = NmToNrm(TthNm)
NmTo2DashSfxTst = A & "__Tst"
End Function

Sub NmTo2DashSfxTst__Tst()
Debug.Assert NmTo2DashSfxTst("Tst__lsdf") = "lsdf__Tst"
Debug.Assert NmTo2DashSfxTst("Tst_lsdf") = "lsdf__Tst"
Debug.Assert NmTo2DashSfxTst("lsdf_Tst") = "lsdf__Tst"
Debug.Assert NmTo2DashSfxTst("lsdf__Tst") = "lsdf__Tst"
End Sub

Function NmToNrm$(Nm$)
Dim A$
A = Replace(Nm, "Tst__", "")
A = Replace(A, "__Tst", "")
A = Replace(A, "_Tst", "")
NmToNrm = Replace(A, "Tst_", "")
End Function

Function NmToPfxTst2Dash$(TthNm$)
Dim A$: A = TthNm
A = RmvPfx(A, "Tst__")
A = RmvPfx(A, "Tst_")
A = RmvSfx(A, "__Tst")
A = RmvSfx(A, "_Tst")
NmToPfxTst2Dash = "Tst__" & A
End Function

Sub NmToPfxTst2Dash__Tst()
Debug.Assert NmToPfxTst2Dash("Tst__lsdf") = "Tst__lsdf"
Debug.Assert NmToPfxTst2Dash("Tst_lsdf") = "Tst__lsdf"
Debug.Assert NmToPfxTst2Dash("lsdf_Tst") = "Tst__lsdf"
Debug.Assert NmToPfxTst2Dash("lsdf__Tst") = "Tst__lsdf"
End Sub

Function NmToTstNm$(Nm$)
NmToTstNm = NmToNrm(Nm) & "__Tst"
End Function
