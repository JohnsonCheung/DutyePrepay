VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSKU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Option Base 0

Private Sub CmdClose_Click()
DoCmd.Close
End Sub

Private Sub CmdSetBusArea_Click()
DoCmd.OpenForm "frmBF"
End Sub

Private Sub Form_Open(Cancel As Integer)
DoCmd.Maximize
End Sub
