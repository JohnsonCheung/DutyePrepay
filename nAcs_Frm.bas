Attribute VB_Name = "nAcs_Frm"
Option Compare Database
Option Explicit
Type FrmOpt
    Som As Boolean
    Frm As Access.Form
End Type

Sub FrmCls(FrmNm$)
DoCmd.Close acForm, FrmNm, acSaveYes
End Sub

Sub FrmSavRec(A As Access.Form)
Stop
End Sub

Function FrmCtlNy(Optional A As Access.Form) As String()
Dim F As Access.Form
    Set F = FrmNz(A)
Dim C As Access.Control
Dim O$()
For Each C In F.Controls
    Push O, TypeName(C) & ":" & C.Name
Next
FrmCtlNy = AySrt(O)
End Function

Function FrmIsOpn(FrmNm$, A As Access.Application) As Boolean
On Error GoTo R
FrmIsOpn = AppaNz(A).Forms(FrmNm).Name
R:
End Function

Function FrmNz(A As Access.Form) As Access.Form
If IsNothing(A) Then
    Set FrmNz = Application.Screen.ActiveForm
Else
    Set FrmNz = A
End If
End Function

Function FrmOpn(FrmNm$, Optional OpnArg, Optional IsDialog As Boolean) As FrmOpt
If IsDialog Then
    DoCmd.OpenForm FrmNm, , , , , acDialog, OpnArg
    Exit Function
End If
DoCmd.OpenForm FrmNm, , , , , , OpnArg
FrmOpn = FrmOpt(Access.Application.Forms(FrmNm))
End Function

Sub FrmOpn__Tst()
Const FrmNm$ = "A"
Dim A As FrmOpt: A = FrmOpn(FrmNm)
FrmCls FrmNm
End Sub

Function FrmOpt(A As Access.Form) As FrmOpt
Set FrmOpt.Frm = A
FrmOpt.Som = True
End Function
