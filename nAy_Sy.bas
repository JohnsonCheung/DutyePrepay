Attribute VB_Name = "nAy_Sy"
Option Compare Database
Option Explicit

Function SyRead(Ft$) As String()
Dim O$()
Dim F%
F = FtOpnInp(Ft)
Dim L$
While Not EOF(F)
    Line Input #F, L
    Push O, L
Wend
SyRead = O
End Function
