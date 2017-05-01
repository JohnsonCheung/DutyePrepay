Attribute VB_Name = "nVb_nNmstr_Nm"
Option Compare Database
Option Explicit

Function NmExpd(Nm, Ay$()) As String()
If StrHas(Nm, "*") Then
    NmExpd = AySelLik(Ay, Nm)
Else
    If AyHas(Ay, Nm) Then
        NmExpd = ApSy(Nm)
    End If
End If
End Function
