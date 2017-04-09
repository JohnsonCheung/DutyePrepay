Attribute VB_Name = "nXls_nRfh_Pt"
Option Compare Database
Option Explicit

Sub PtRfh(A As PivotTable)
On Error Resume Next
A.RefreshTable
End Sub
