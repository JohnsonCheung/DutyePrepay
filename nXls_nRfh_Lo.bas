Attribute VB_Name = "nXls_nRfh_Lo"
Option Compare Database
Option Explicit

Function LoRfh(A As ListObject)
On Error Resume Next
QtRfh A.QueryTable
End Function
