VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPermitD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private x_Permit&

Private Sub Amt_AfterUpdate()
If Nz(Me.Qty.Value, 0) = 0 Then Me.Rate.Value = 0: Exit Sub
Me.Rate.Value = Nz(Me.Amt.Value, 0) / Nz(Me.Qty.Value, 0)
End Sub

Private Sub Amt_BeforeUpdate(Cancel As Integer)
If IsNull(Me.Amt.Value) Then Cancel = True: MsgBox "Cannot be blank", vbCritical
If Me.Amt.Value <= 0 Then Cancel = True: MsgBox "Cannot be -ve or zero", vbCritical
End Sub

Private Sub BchNo_AfterUpdate()
ZCrtSkuB
End Sub

Private Sub CmdClose_Click()
DoCmd.Close
End Sub

Private Sub CmdDlt_Click()
AppaSavRec
If IsNull(Me.Sku.Value) Then Exit Sub
SqlRun Fmt_Str("Delete From PermitD where Permit={0} and Sku='{1}'", x_Permit, Me.Sku.Value)
Me.Requery
End Sub

Private Sub CmdGenFx_Click()
AppaSavRec
If IsNull(Me.Permit.Value) Then Exit Sub
FrmPermitCmdGenFx Me.Permit.Value
End Sub

Private Sub Form_AfterUpdate()
ZRfhTot
UpdSKURate Me.Sku.Value, Me.Rate.Value
Me.yCheck.Requery
Me.yChkZHT0vsInput.Requery
Me.yDutyRateB.Requery
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
Me.Permit.Value = x_Permit
Me.SeqNo.Value = ZNxtSeqNo(x_Permit)
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
Me.DteUpd.Value = Now
Qty_BeforeUpdate Cancel: If Cancel Then Me.Qty.SetFocus: Exit Sub
Amt_BeforeUpdate Cancel: If Cancel Then Me.Amt.SetFocus: Exit Sub
Rate_BeforeUpdate Cancel: If Cancel Then Me.Rate.SetFocus: Exit Sub
End Sub

Private Sub Form_Close()
Form_Close_1UpdPermitTot
TblSkuBBchRateErOpn
End Sub

Private Sub Form_Close_1UpdPermitTot()
Dim mNSku%, mQty&, mTot@
With CurrentDb.OpenRecordset("SELECT Count(*) AS NSku, Sum(x.Qty) AS Qty, Sum(Amt) AS Tot FROM PermitD x WHERE Permit=" & x_Permit)
    If Not .EOF Then
        mNSku = Nz(!NSku, 0)
        mQty = Nz(!Qty, 0)
        mTot = Nz(!Tot, 0)
        .Close
    End If
End With
With CurrentDb.OpenRecordset("SELECT * from Permit where Permit=" & x_Permit)
    If Not .EOF Then
        If !NSku <> mNSku Or !Qty <> mQty Or !Tot <> mTot Then
            .Edit
            !NSku = mNSku
            !Qty = mQty
            !Tot = mTot
            !DteUpd = Now()
            .Update
        End If
    End If
    .Close
End With
End Sub

Private Sub Form_Open(Cancel As Integer)
x_Permit = Me.OpenArgs
Me.xPermit.Value = x_Permit
Me.xPermitDate.Value = TblPermitDate(x_Permit)
Me.xPermitNo.Value = TblPermitNo(x_Permit)
TblfrmPermitDBld x_Permit
Me.RecordSource = "Select * from PermitD x inner join frmPermitD a on a.PermitD=x.PermitD where Permit=" & x_Permit
ZRfhTot
Requery
Refresh
DoCmd.Maximize
End Sub

Private Sub Qty_AfterUpdate()
ZRfhAmt
End Sub

Private Sub Qty_BeforeUpdate(Cancel As Integer)
If IsNull(Me.Qty.Value) Then Cancel = True: MsgBox "Cannot be blank", vbCritical
If Me.Qty.Value <= 0 Then Cancel = True: MsgBox "Cannot be -ve or zero", vbCritical
End Sub

Private Sub Rate_AfterUpdate()
ZRfhAmt
ZCrtSkuB
End Sub

Private Sub Rate_BeforeUpdate(Cancel As Integer)
If IsNull(Me.Rate.Value) Then Cancel = True: MsgBox "Cannot be blank", vbCritical
If Me.Rate.Value <= 0 Then Cancel = True: MsgBox "Cannot be -ve or zero", vbCritical
End Sub

Private Sub SKU_AfterUpdate()
ZCrtSkuB
End Sub

Private Sub SKU_BeforeUpdate(Cancel As Integer)
Dim R As OptCur
R = TblSkuRatOpt(Me.Sku.Value)
If Not R.Som Then Cancel = True
Me.Rate.Value = R.Cur
End Sub

Private Sub UpdSKURate(pSku$, pRate@)
With CurrentDb.OpenRecordset("Select * from SKU where Sku='" & pSku & "'")
    If .EOF Then
        .AddNew
        !Sku = pSku
        !DutyRate = pRate
        .Update
    Else
        If !DutyRate <> pRate Then
            .Edit
            !DutyRate = pRate
            !DteUpd = Now()
            .Update
        End If
    End If
    .Close
End With
End Sub

Private Sub ZCrtSkuB()
Dim mSku, mBchNo, MRate
mSku = Me.Sku.Value
mBchNo = Me.BchNo.Value
MRate = Me.Rate.Value
If IsNull(mSku) Or IsNull(mBchNo) Or IsNull(MRate) Then Exit Sub
With CurrentDb.OpenRecordset(Fmt_Str("Select * from SkuB where Sku='{0}' and BchNo='{1}'", mSku, mBchNo))
    If .EOF Then
        .AddNew
        !Sku = mSku
        !BchNo = mBchNo
        !DutyRateB = MRate
        .Update
    Else
        If !DutyRateB <> MRate Then
            .Edit
            !DutyRateB = MRate
            .Update
        End If
    End If
    .Close
End With
End Sub

Private Function ZNxtSeqNo%(PermitId&)
With CurrentDb.OpenRecordset("Select Max(SeqNo) from PermitD where Permit=" & PermitId)
    ZNxtSeqNo = 10 + Nz(.Fields(0).Value, 0)
    .Close
End With
End Function

Private Sub ZRfhAmt()
Me.Amt.Value = Nz(Me.Rate.Value, 0) * Nz(Me.Qty.Value, 0)
End Sub

Private Sub ZRfhTot()
With CurrentDb.OpenRecordset("Select Count(*) as NSku, Sum(d.Qty) as Qty, Sum(Amt) as Tot from PermitD d where Permit=" & x_Permit)
    If .EOF Then
        Me.xQty.Value = 0
        Me.xNSku.Value = 0
        Me.xTot.Value = 0
    Else
        Me.xQty.Value = !Qty
        Me.xNSku.Value = !NSku
        Me.xTot.Value = !Tot
    End If
    .Close
End With
End Sub
