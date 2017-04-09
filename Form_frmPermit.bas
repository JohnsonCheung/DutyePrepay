VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPermit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Option Base 0
Private xGLAc$, xGLAcName$, xByUsr$, xBankCode$
Public X_CurPermitNo$

Function CurPermitNo$()
CurPermitNo = X_CurPermitNo
End Function

Function IsCurPermitNo(PermitNo$) As Boolean
IsCurPermitNo = X_CurPermitNo = PermitNo
End Function

Function SetCurPermitNo$(PermitNo$)
X_CurPermitNo = PermitNo
End Function

Private Sub CmdClose_Click()
DoCmd.Close
End Sub

Private Sub CmdDelete_Click()
End Sub

Private Sub CmdDlt_Click()
AppaSavRec
If IsNull(Me.PermitNo.Value) Then Exit Sub
Dim A$: A = UCase(InputBox("INPUT [YES] to delete permitNo-" & Me.PermitNo.Value, "Delete Permit"))
If A <> "YES" Then Exit Sub
Dim P$: P = Me.PermitNo.Value
SqlRunQQ "DELETE FROM PERMITD WHERE Permit IN (SELECT PERMIT FROM PERMIT WHERE PERMITNO='?')", P
SqlRunQQ "DELETE FROM PERMIT WHERE PermitNo='?'", P
Me.Requery
End Sub

Private Sub CmdEdt_Click()
AppaSavRec
If IsNull(Me.Permit.Value) Then Exit Sub
DoCmd.OpenForm "frmPermitD", acNormal, DataMode:=acFormEdit, OpenArgs:=Me.Permit.Value
End Sub

Private Sub CmdExpList_Click()
FrmPermitCmdExpList
End Sub

Private Sub CmdGenFx_Click()
AppaSavRec
If IsNull(Me.Permit.Value) Then MsgBox "Please enter Permit#": Exit Sub
FrmPermitCmdGenFx Me.Permit.Value
End Sub

Private Sub CmdImp_Click()
FrmPermitCmdImp CStr(Me.PermitNo.Value)
Me.Requery
End Sub

Private Sub CmdOpnChqFdr_Click()
PthBrw FbCurPth & "Cheque Request\"
End Sub

Private Sub CmdOpnImpFdr_Click()
PthBrw ImpPermitFdr
End Sub

Private Sub CmdReadMeV6_Click()
FrmPermitCmdReadMeV6
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
Me.GLAc.Value = xGLAc
Me.GLAcName.Value = xGLAcName
Me.ByUsr.Value = xByUsr
Me.BankCode.Value = xBankCode
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
Me.DteUpd.Value = Now
End Sub

Private Sub Form_Current()
Me.CmdImp.Enabled = Me.CanImp
SetCurPermitNo CStr(Me.PermitNo.Value)
Me.Recalc
End Sub

Private Sub Form_Open(Cancel As Integer)
TblPermitRfhFldCanImp
DoCmd.Maximize
With CurrentDb.OpenRecordset("Select * from Default")
    xGLAc = !GLAc
    xGLAcName = !GLAcName
    xByUsr = !ByUsr
    xBankCode = !BankCode
    .Close
End With
TblSkuBRfh
DoCmd.RunCommand acCmdRemoveFilterSort
End Sub

Private Sub PermitDate_BeforeUpdate(Cancel As Integer)
If IsNull(Me.PermitDate.Value) Then MsgBox "Cannot be blank": Cancel = True: Exit Sub
If Me.PermitDate.Value < #1/1/2010# Then MsgBox "Cannot less then 2010/01/01": Cancel = True: Exit Sub
If Me.PermitDate.Value > #1/1/2050# Then MsgBox "Cannot greater then 2050/01/01": Cancel = True: Exit Sub
If VBA.Year(Me.PermitDate.Value) <> VBA.Year(Date) Then If MsgBox("The year is not current year, is it OK?", vbOKCancel) = vbCancel Then Cancel = True
End Sub

Private Sub PostDate_BeforeUpdate(Cancel As Integer)
If IsNull(Me.PostDate.Value) Then MsgBox "Cannot be blank": Cancel = True: Exit Sub
If Me.PostDate.Value < #1/1/2000# Then MsgBox "Cannot less then 2010/01/01": Cancel = True: Exit Sub
If Me.PostDate.Value > #1/1/2050# Then MsgBox "Cannot greater then 2050/01/01": Cancel = True: Exit Sub
If VBA.Year(Me.PostDate.Value) <> VBA.Year(Date) Then If MsgBox("The year is not current year, is it OK?", vbOKCancel) = vbCancel Then Cancel = True
End Sub
