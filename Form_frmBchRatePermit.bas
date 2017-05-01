VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmBchRatePermit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Option Base 0

Private Sub BchNo_BeforeUpdate(Cancel As Integer)
Dim MRate@
If IsNull(Me.Sku.Value) Then MsgBox "SKU is empty": Cancel = True: Exit Sub
Cancel = BchNo_BeforeUpdate_1BchNo(Me.Sku.Value, Me.BchNo.Value): If Cancel Then Exit Sub
Cancel = VdtBchNo(Nz(Me.BchNo.Value), Nz(Me.Sku.Value, ""), Nz(Me.Rate.Value, 0), MRate)
End Sub

Private Function BchNo_BeforeUpdate_1BchNo(pSku$, pBchNo) As Boolean
'Aim: If pBchNo is null, return false for no error
'     If pSku+pBchNo not found in frmBchRateOH, prompt message and return true for error
'Assume: pSku is non-blank
If IsNull(pBchNo) Then Exit Function
If Trim(pBchNo) = "" Then Exit Function
With CurrentDb.OpenRecordset(Fmt("Select * from frmBchRateOH where SKU='{0}' and BchNo='{1}'", pSku, pBchNo))
    If .EOF Then
        .Close
        MsgBox "The no such batch on hand!", vbCritical
        BchNo_BeforeUpdate_1BchNo = True
        Exit Function
    End If
    .Close
End With
End Function

Private Function VdtBchNo(BchNo$, Sku$, Rate@, ByRef ORate@) As Boolean
'Aim: Validate BchNo:
'     - Return false for no error for BchNo=''
'     - Return true for error for Sku='' or Rate=0
'     - If there is a record in SkuB return false. Set oRate=SkuB->Rate
'     - If there is no recor in SkuB return false. Add one record to SkuB, set oRate=Rate
If BchNo = "" Then Exit Function
If Sku = "" Then MsgBox "SKU is blank": GoTo E
If Rate = 0 Then MsgBox "Rate is zero": GoTo E
With CurrentDb.OpenRecordset(Fmt("Select * from SkuB where Sku='{0}' and BchNo='{1}'", Sku, BchNo))
    If .EOF Then
        ORate = Rate
        .AddNew
        !DutyRateB = Rate
        !Sku = Sku
        !BchNo = BchNo
        .Update
        .Close
        Exit Function
    End If
    ORate = !DutyRateB
    .Close
End With
Exit Function
E: VdtBchNo = True
End Function
