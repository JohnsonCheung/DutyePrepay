Attribute VB_Name = "nAy_Ly"
Option Compare Database
Option Explicit

Function LyJn$(Ay)
Dim A(): A = AyExpandAy(Ay)
LyJn = AyJn(Ay, vbCrLf)
End Function

