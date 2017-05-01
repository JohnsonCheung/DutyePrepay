Attribute VB_Name = "mFrmPermitD_g"
Option Compare Database
Option Explicit

Function gChkRate$(Rate, RateDuty)     ' Used by frmPermitD
If IsNull(RateDuty) Then gChkRate = "No Rate!": Exit Function  ' RateDuty is ZHT0 rate
If Nz(Rate, 0) = 0 Then gChkRate = "---": Exit Function
If Rate > RateDuty Then
    If (Rate - RateDuty) / Rate > 0.1 Then gChkRate = "Too low": Exit Function
Else
    If (RateDuty - Rate) / Rate > 0.1 Then gChkRate = "Too high": Exit Function
End If
End Function

Function gTblSkuBRateB@(Sku, BchNo)     ' Used by frmPermitD
If IsNull(Sku) Then Exit Function         ' Rate     is PermitD->Rate which is user input
If IsNull(BchNo) Then Exit Function     ' RateDuty is ZHT0 rate
With CurrentDb.OpenRecordset(Fmt("Select DutyRateB from SkuB where Sku='{0}' and BchNo='{1}'", Sku, BchNo))
    If .EOF Then .Close: Exit Function
    gTblSkuBRateB = Nz(.Fields(0).Value, 0)
    .Close
End With
End Function
