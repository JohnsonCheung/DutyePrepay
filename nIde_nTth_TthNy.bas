Attribute VB_Name = "nIde_nTth_TthNy"
Option Compare Database
Option Explicit

Function TthNy_Md(Optional A As CodeModule) As String()
TthNy_Md = TthNy_MdXXX(A)
End Function

Function TthNy_MdPri(Optional A As CodeModule) As String()
TthNy_MdPri = TthNy_MdXXX(A)
End Function

Function TthNy_MdPriPfx(Optional A As CodeModule) As String()
TthNy_MdPriPfx = TthNy_MdXXX(A, "PriPfx")
End Function

Function TthNy_MdPriSfx(Optional A As CodeModule) As String()
TthNy_MdPriSfx = TthNy_MdXXX(A, "PriSfx")
End Function

Function TthNy_MdPub(Optional A As CodeModule) As String()
TthNy_MdPub = TthNy_MdXXX(A, "Pub")
End Function

Sub TthNy_MdPub__Tst()
AyBrw TthNy_MdPub
End Sub

Function TthNy_MdPubPfx(Optional A As CodeModule) As String()
TthNy_MdPubPfx = TthNy_MdXXX(A, "PubPfx")
End Function

Function TthNy_MdPubSfx(Optional A As CodeModule) As String()
TthNy_MdPubSfx = TthNy_MdXXX(A, "PubSfx")
End Function

Function TthNy_Pj(Optional A As vbproject) As String()
TthNy_Pj = TthNy_PjXXX(A)
End Function

Sub TthNy_Pj__Tst()
DrAyBrw AyBrk(TthNy_Pj, ".")
End Sub

Function TthNy_PjPri(Optional A As CodeModule) As String()
TthNy_PjPri = TthNy_PjXXX(A, "Pri")
End Function

Sub TthNy_PjPri__Tst()
AyBrw TthNy_PjPri
End Sub

Function TthNy_PjPriPfx(Optional A As CodeModule) As String()
TthNy_PjPriPfx = TthNy_PjXXX(A, "PriPfx")
End Function

Function TthNy_PjPriSfx(Optional A As CodeModule) As String()
TthNy_PjPriSfx = TthNy_PjXXX(A, "PriSfx")
End Function

Function TthNy_PjPub(Optional A As CodeModule) As String()
TthNy_PjPub = TthNy_PjXXX(A, "Pub")
End Function

Function TthNy_PjPubPfx(Optional A As CodeModule) As String()
TthNy_PjPubPfx = TthNy_PjXXX(A, "PubPfx")
End Function

Function TthNy_PjPubSfx(Optional A As CodeModule) As String()
TthNy_PjPubSfx = TthNy_PjXXX(A, "PubSfx")
End Function

Sub TthRen_PjTo2DashSfxTst__Tst()
TthRen_PjTo2DashSfxTst
End Sub

Sub TthRen_PjToPfxTst2Dash__Tst()
TthRen_PjToPfxTst2Dash
End Sub

Sub TthRen_To2DashSfxTstMd__Tst()
TthRen_MdTo2DashSfxTst Md("nDta_Dt")
End Sub

Private Function TthNy_MdXXX(A As CodeModule, Optional FctSfx$) As String()
Dim B() As MthBrk: B = MthBrkAy(A)
If MthBrkIsEmptyAy(B) Then Exit Function
Dim J%, O$()
For J = 0 To UBound(B)
    Select Case FctSfx
    Case "":       If MthBrkIsTth(B(J)) Then Push O, B(J).Nm
    Case "Pub":    If MthBrkIsTth_Pub(B(J)) Then Push O, B(J).Nm
    Case "Pri":    If MthBrkIsTth_Pri(B(J)) Then Push O, B(J).Nm
    Case "PubSfx": If MthBrkIsTth_PubSfx(B(J)) Then Push O, B(J).Nm
    Case "PubPfx": If MthBrkIsTth_PubPfx(B(J)) Then Push O, B(J).Nm
    Case "PriSfx": If MthBrkIsTth_PriSfx(B(J)) Then Push O, B(J).Nm
    Case "PriPfx": If MthBrkIsTth_PriPfx(B(J)) Then Push O, B(J).Nm
    Case Else: Er "Invalid {FctSfx}", FctSfx
    End Select
Next
TthNy_MdXXX = O
End Function

Private Function TthNy_PjXXX(A As vbproject, Optional FctSfx$) As String()
Dim B() As CodeModule: B = PjMdAy(A)
Dim J%, O$(), Ay$(), DrAy()
For J = 0 To UBound(B)
    Ay = Run("TthNy_Md" & FctSfx, B(J))
    Ay = AyAddPfx(Ay, MdNm(B(J)) & ".")
    PushAy O, Ay
Next
TthNy_PjXXX = O
End Function
