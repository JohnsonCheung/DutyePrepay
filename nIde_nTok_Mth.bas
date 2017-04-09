Attribute VB_Name = "nIde_nTok_Mth"
Option Compare Database
Option Explicit

Function MthTokNy(MthNm$, Optional A As CodeModule) As String()
Dim Ly$(): Ly = MthLy(MthNm, A)
Dim Ay1$(): Ay1 = LyTokAy(MthLy(MthNm, A))
Dim Ay2$(): Ay2 = MthLyDefTokNy(Ly)
MthTokNy = AyMinus(Ay1, Ay2)
End Function

Sub MthTokNy__Tst()
AyDmp MthTokNy("MthTokNy", Md("nIde_nTok_Mth"))
End Sub
