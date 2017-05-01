Attribute VB_Name = "nVb_LpAp"
Option Compare Database
Option Explicit

Function LpApToStr$(SepChr$, Lp$, ParamArray Ap())
Dim A$(): A = Split(Lp, ",")
Dim K, J%, O$()
For Each K In A
    Push O, K & "=[" & VarToStr(Ap(J)) & "]"
    J = J + 1
Next
LpApToStr = Jn(O, SepChr)
End Function

Function LpApToStr__Tst()
Debug.Print LpApToStr(vbLf, "aa,bb,,C", 1, 2, , 1)
End Function
