Attribute VB_Name = "nXls_nRfh_Cht"
Option Compare Database
Option Explicit

Sub ChtRfh(A As Chart)
On Error Resume Next
If IsNothing(A.PivotLayout) Then Exit Sub
PtRfh A.PivotLayout.PivotTable
End Sub
