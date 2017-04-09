VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmBchRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Option Base 0
Private xYY As Byte, xMM As Byte, xDD As Byte

Private Sub CmdAddToDummyPermit_Click()
If IsNull(Me.IsNoAssign.Value) Then Exit Sub
If Not Me.IsNoAssign.Value Then MsgBox "SKU[" & Me.Sku.Value & "] does not need to add to dummy permit", vbInformation: Exit Sub
FrmBchRateCmdAddToDmyPermit CStr(Me.Sku.Value)
Requery
End Sub

Private Sub CmdAssign_Click()
FrmBchRateCmdAssign xYY, xMM, xDD
Me.Requery
Me.Refresh
End Sub

Private Sub CmdClose_Click()
DoCmd.Close
End Sub

Private Sub CmdReadMe_Click()
MsgBox "Each latest inventory on hand SKU + batch# of taxable item in company 8600 at location Consignment & TaxPaid should assign to a permit line so that the inventory Duty-Rate of the inventory can be determined." & vbLf & vbLf & _
"Click [Assign Batch# to permit] will automatically assign batch# to permit." & vbLf & vbLf & _
"However some SKU may not able find any permit to assign.  In this case, [Cannot Assign] will be checked.  Click [Assign to dummy permit].  User can input the a Tax Duty in this dummy permit."
End Sub

Private Sub Form_Close()
TblSkuBBchRateErOpn
End Sub

Private Sub Form_Open(Cancel As Integer)
'Aim: Set xDteOH
'     Build "frmBchRateOH = SKU OH from
DoCmd.Maximize
With TblOHMaxYMD
    xYY = .Y
    xMM = .M
    xDD = .M
End With
Me.xDteOH.Value = "20" & Format(xYY, "00") & "-" & Format(xMM, "00") & "-" & Format(xDD, "00")
TblSkuBRfh
zBldRecordSourceTable xYY, xMM, xDD
End Sub
