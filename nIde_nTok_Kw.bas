Attribute VB_Name = "nIde_nTok_Kw"
Option Compare Database
Option Explicit

Function KwAy() As String()
Static X$()
If AyIsEmpty(X) Then X = Split("Compare Print Type With Property Get Set Let Optional Select Case Function Sub String As Integer Long Short If Then Else End For To Next On Error Goto While Not Wend Option Explicit")
KwAy = X
End Function

Private Sub KwAy__Tst()
AyBrw KwAy
End Sub
