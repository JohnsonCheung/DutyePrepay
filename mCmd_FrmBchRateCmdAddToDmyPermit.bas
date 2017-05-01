Attribute VB_Name = "mCmd_FrmBchRateCmdAddToDmyPermit"
Option Compare Database
Option Explicit

Sub FrmBchRateCmdAddToDmyPermit(Sku)
Cmd_AddToDmyPermit_Click_1Add Sku
Cmd_AddToDmyPermit_Click_2Reset Sku
Cmd_AddToDmyPermit_Click_3UpdPermitThreeSum
End Sub

Private Sub Cmd_AddToDmyPermit_Click_1Add(Sku)
Cmd_AddToDmyPermit_Click_1Add_1Permit
Cmd_AddToDmyPermit_Click_1Add_2PermitD Sku
End Sub

Private Sub Cmd_AddToDmyPermit_Click_1Add_1Permit()
With CurrentDb.OpenRecordset("Select * from Permit where Permit=1")
    If .EOF Then
        .Close
        SqlRun "Insert into Permit (Permit,PermitNo,PermitDate,PostDate,GLAc,GLAcName,BankCode,ByUsr) values (1,'--Dmy--',#2000/1/1#,#2000/1/1#,'-','-','-','-')"
        Exit Sub
    End If
    .Close
End With
End Sub

Private Sub Cmd_AddToDmyPermit_Click_1Add_2PermitD(Sku)
With CurrentDb.OpenRecordset(Fmt("Select * from frmBchRateOH where Sku='{0}' and IsNoAssign", Sku))
    While Not .EOF
        Cmd_AddToDmyPermit_Click_1Add_2PermitD_1Ins Sku, !BchNo, !OH
        .MoveNext
    Wend
    .Close
End With
End Sub

Private Sub Cmd_AddToDmyPermit_Click_1Add_2PermitD_1Ins(Sku, BchNo$, Qty&)
Dim mSeqNo%: mSeqNo = Nz(SqlV("Select Max(SeqNo) from PermitD where Permit=1"), 0) + 10
Dim MRate@: MRate = Nz(SqlV(Fmt("Select DutyRateBott from frmBchRate where Sku='{0}'", Sku)), 0)
Dim mAmt@: mAmt = Qty * MRate
With CurrentDb.TableDefs("PermitD").OpenRecordset
    .AddNew
    !Permit = 1
    !Sku = Sku
    !SeqNo = mSeqNo
    !Qty = Qty
    !BchNo = BchNo
    !Rate = MRate
    !Amt = mAmt
    .Update
    .Close
End With
End Sub

Private Sub Cmd_AddToDmyPermit_Click_2Reset(Sku)
SqlRun Fmt("Update frmBchRateOH set IsNoAssign=False where IsNoAssign and SKU='{0}'", Sku)
SqlRun Fmt("Update frmBchRate    set IsNoAssign=False where IsNoAssign and SKU='{0}'", Sku)
End Sub

Private Sub Cmd_AddToDmyPermit_Click_3UpdPermitThreeSum()
Dim mNSku%: mNSku = Nz(SqlV("Select Count(*) from PermitD where Permit=1"), 0)
Dim mQty&: mQty = Nz(SqlV("Select Sum(Qty) from PermitD where Permit=1"), 0)
Dim mTot@: mTot = Nz(SqlV("Select Sum(Amt) from PermitD where Permit=1"), 0)
With CurrentDb.OpenRecordset("Select * from Permit where Permit=1")
    .Edit
    !DteUpd = Now()
    !NSku = mNSku
    !Qty = mQty
    !Tot = mTot
    .Update
End With
End Sub
