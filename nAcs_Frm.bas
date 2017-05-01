Attribute VB_Name = "nAcs_Frm"
Option Compare Database
Option Explicit
Type FrmOpt
    Som As Boolean
    Frm As Access.Form
End Type

Sub FrmCls(A As Access.Form)
DoCmd.Close acForm, A.Name, acSaveYes
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

Function FrmHasCtl(A As Access.Form, CtlNm$) As Boolean
On Error Resume Next
FrmHasCtl = A.Controls(CtlNm).Name = CtlNm: Exit Function
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

Function FrmOpn(FrmNm$, Optional OpnArg, Optional IsDialog As Boolean) As Access.Form
If IsDialog Then
    DoCmd.OpenForm FrmNm, , , , , acDialog, OpnArg
    Exit Function
End If
DoCmd.OpenForm FrmNm, , , , , , OpnArg
Set FrmOpn = Access.Application.Forms(FrmNm)
End Function

Sub FrmOpn__Tst()
Const FrmNm$ = "A"
Dim A As Access.Form: Set A = FrmOpn(FrmNm)
FrmCls A
End Sub

Function FrmOpnOpt(FrmNm$, Optional OpnArg, Optional IsDialog As Boolean) As FrmOpt
If IsDialog Then
    DoCmd.OpenForm FrmNm, , , , , acDialog, OpnArg
    Exit Function
End If
DoCmd.OpenForm FrmNm, , , , , , OpnArg
FrmOpnOpt = FrmOpt(Appa.Forms(FrmNm))
End Function

Function FrmOpt(A As Access.Form) As FrmOpt
Set FrmOpt.Frm = A
FrmOpt.Som = True
End Function

Sub FrmSavRec(A As Access.Form)
Stop
End Sub

Sub FrmSetChdLnk(A As Access.Form, SubFrmNm$, FldDic As Dictionary)
With A.Controls(SubFrmNm)
    .LinkMasterFields = DicSemiColonKeyStr(FldDic)
    .LinkChildFields = DicSemiColonValStr(FldDic)
End With
End Sub

Sub FrmSetLck(A As Access.Form, Lck As Boolean, Optional AlwAdd As Boolean, Optional AlwDlt As Boolean)
'Aim: Set all controls in {A} as lock
Dim Ctl As Access.Control: For Each Ctl In A.Controls
    If Not Visible Then GoTo Nxt
    Dim mLck As Boolean: If Ctl.Tag = "Edt" Then mLck = Lck Else mLck = True
    Select Case TypeName(Ctl)
    Case "Label":    LblSetLck Ctl, Lck
    Case "TextBox":  TBoxSetLck Ctl, Lck
    Case "Check":    ChkBSetLck Ctl, Lck
    Case "ComboBox": CBoxSetLck Ctl, Lck
    End Select
Nxt:
Next
With A
    If Lck Then
        .AllowEdits = False
        .AllowAdditions = False
        .AllowDeletions = False
    Else
        .AllowEdits = True
        .AllowAdditions = AlwAdd
        .AllowDeletions = AlwDlt
    End If
End With
A.Repaint
End Sub

Function FrmSetLck__Tst()
Const cNmFrm$ = "frmIIC_Tst"
Dim F As Access.Form: Set F = FrmOpn(cNmFrm)
FrmSetLck F, False
FrmSetLck F, True
Stop
FrmSetLck F, False
Stop
FrmCls F
End Function

Function FrmToStr$(A As Access.Form)
On Error GoTo R
FrmToStr = A.Name
Exit Function
R: FrmToStr$ = ErStr("FrmToStr")
End Function
