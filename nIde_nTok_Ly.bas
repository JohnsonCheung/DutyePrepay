Attribute VB_Name = "nIde_nTok_Ly"
Option Compare Database
Option Explicit

Function LyTokAy(Ly$()) As String()
Dim J&
Dim O$()
For J = 0 To UB(Ly)
    PushAyNoDup O, LinTokAy(Ly(J))
Next
LyTokAy = O
End Function
