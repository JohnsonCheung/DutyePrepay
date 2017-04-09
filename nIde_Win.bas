Attribute VB_Name = "nIde_Win"
Option Compare Database
Option Explicit

Sub WinClsAll()
Dim I As VBIDE.Window
For Each I In Application.Vbe.Windows
    If I.Type = vbext_wt_Browser Then GoTo Nxt
    If I.Type = vbext_wt_Immediate Then GoTo Nxt
    I.Close
Nxt:
Next
WinEns WinObj
WinEns WinImm
Dim W As VBIDE.Window: Set W = WinObj
If Not W.Visible Then W.Visible = True
PjSav
End Sub

Sub WinEns(A As VBIDE.Window)
If Not A.Visible Then A.Visible = True
End Sub

Function WinImm() As VBIDE.Window
Set WinImm = Application.Vbe.Windows("Immediate")
End Function

Function WinObj() As VBIDE.Window
Set WinObj = Application.Vbe.Windows("Object Browser")
End Function

