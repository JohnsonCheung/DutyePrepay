Attribute VB_Name = "nXls_nRfh_ChtObj"
Option Compare Database
Option Explicit

Sub ChtObjRfh(A As ChartObject)
On Error Resume Next
If IsNothing(A.Chart) Then Exit Sub
ChtRfh A.Chart
End Sub
