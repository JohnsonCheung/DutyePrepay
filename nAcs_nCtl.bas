Attribute VB_Name = "nAcs_nCtl"
Option Compare Database
Option Explicit

Sub CBoxSetEna(A As Access.ComboBox, Ena As Boolean)
A.Enabled = Ena
A.ForeColor = IIf(Ena, 0, 255)
End Sub

Sub CBoxSetLck(A As Access.ComboBox, Lck As Boolean)
A.Locked = Lck
A.ForeColor = 0
A.TabStop = Not Lck
A.ForeColor = IIf(Lck, 255, 0)
End Sub

Sub ChkBSetColr(A As Access.CheckBox, Ena As Boolean)
If Ena Then
    A.BorderColor = 65280
Else
    A.BorderColor = 13209
End If
End Sub

Sub ChkBSetEna(A As Access.CheckBox, Ena As Boolean)
A.Enabled = Ena
A.BorderColor = IIf(Ena, 65280, 13209)
End Sub

Sub ChkBSetLck(A As Access.CheckBox, Lck As Boolean)
A.Locked = Lck
A.BorderColor = IIf(Lck, 13209, 65280)
A.TabStop = Not Lck
End Sub

Sub CtlAySetLayout(A As Access.Form, pLnCtl$, Optional L! = -1, Optional T! = -1, Optional W! = -1, Optional H! = -1)
Dim mAnCtl$(): mAnCtl = Split(pLnCtl, CtComma)
Dim J%: For J = 0 To Sz(mAnCtl) - 1
    Dim C As Access.Control: If Fnd_Ctl(C, A, mAnCtl(J)) Then GoTo Nxt
    CtlSetLayout C, L, T, W, H
Nxt:
Next
End Sub

Function CtlLbl(A As Access.Control) As Access.Label
Dim F As Access.Form: Set F = A.Parent
Dim Nm$
Nm = A.Name & "_Lbl"
If FrmHasCtl(F, Nm) Then Set CtlLbl = A.Controls(Nm)
End Function

Sub CtlSetLayout(A As Access.Control, Optional L! = -1, Optional T! = -1, Optional W! = -1, Optional H! = -1)
With A
    If T >= 0 Then .Top = T
    If L >= 0 Then .Left = L
    If H >= 0 Then .Height = H
    If W >= 0 Then .Width = W
End With
End Sub

Sub CtlSetPrp(A As Access.Control, PrpNm$, V)
A.Properties(PrpNm).Value = V
End Sub

Sub CtlSetPrpInFrm(A As Access.Form, TagSubStr$, PrpNm$, V)
Dim I As Access.Control
For Each I In A.Controls
    If InStr(I.Tag, TagSubStr) > 0 Then CtlSetPrp I, PrpNm, V
Next
End Sub

Sub CtlSetVis(A As Form, TagSubStr$, Vis As Boolean)
Dim I As Control
For Each I In A.Controls
    If InStr(I.Tag, TagSubStr) Then If I.Visible <> Vis Then I.Visible = Vis
Next
End Sub

Sub EdtCtlSetEna(A As Access.Form, Ena As Boolean)
Dim Clt As Access.Control
For Each Clt In A.Controls
    If Clt.Tag = "Edt" Then
        Dim mNmTyp$: mNmTyp = TypeName(Clt)
        Select Case mNmTyp
        Case "TextBox":  TBoxSetEna Clt, Ena
        Case "Check":    ChkBSetEna Clt, Ena
        Case "ComboBox": CBoxSetEna Clt, Ena
        End Select
    End If
Next
End Sub

Sub LblSetColr(A As Label, Ena As Boolean)
With A
    If Ena Then
        .BackColor = 65280
        .ForeColor = 0
    Else
        .BackColor = 13209
        .ForeColor = 16777215
    End If
End With
End Sub

Sub LblSetLck(A As Access.Label, Lck As Boolean)
A.ForeColor = IIf(Lck, 16777215, 0)
A.BackColor = IIf(Lck, 13209, 65280)
End Sub

Function LblSetLck__Tst()
Const cNmFrm$ = "frmIIC_Tst"
Dim F As Access.Form: Set F = FrmOpn(cNmFrm)
Dim L As Access.Label: Set L = F.Controls("ICGL_Label")
LblSetLck L, False
End Function

Sub TBoxSetEna(A As Access.TextBox, Ena As Boolean)
A.Enabled = Ena
A.ForeColor = IIf(Ena, 0, 255)
End Sub

Sub TBoxSetLck(A As Access.TextBox, Lck As Boolean)
A.Locked = Lck
A.ForeColor = 0
A.TabStop = Not Lck
Dim L As Access.Label: Set L = CtlLbl(A)
If Not IsNothing(L) Then LblSetLck L, Lck
End Sub

Sub TBoxSetLck__Tst()
Const cNmFrm$ = "frmIIC_Tst"
Dim F As Access.Form: Set F = FrmOpn(cNmFrm)
Dim T As Access.TextBox: Set T = F.Controls("ICGL")
TBoxSetLck T, False
End Sub
