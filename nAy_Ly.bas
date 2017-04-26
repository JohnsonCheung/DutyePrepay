Attribute VB_Name = "nAy_Ly"
Option Compare Database
Option Explicit

Function LyJn$(Ay)
Dim A(): A = AyExpdAy(Ay)
LyJn = AyJn(Ay, vbCrLf)
End Function

