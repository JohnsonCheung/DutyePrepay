Attribute VB_Name = "nDta_nScl_DrAy"
Option Compare Database
Option Explicit

Function DrAyNewScLy(ScLy$()) As Variant()
Dim U&: U = UB(ScLy)
Dim O(): ReSz O, U
Dim J&
For J = 0 To U
    O(J) = DrNewScl(ScLy(J))
Next
DrAyNewScLy = O
End Function

