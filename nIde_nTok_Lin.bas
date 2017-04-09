Attribute VB_Name = "nIde_nTok_Lin"
Option Compare Database
Option Explicit

Function LinTokAy(Lin) As String()
Dim L$: L = LinRmvRmk(Lin)
LinTokAy = AyMinus(AyRmvDup(LvsSplit(RplPun(RmvStrTok(L)))), KwAy)
End Function

