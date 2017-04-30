Attribute VB_Name = "nXls_nObj_nBdr_nDo_Bdr"
Option Compare Database
Option Explicit

Sub BdrSet_Continuous_Medium(A As Border)
A.LineStyle = XlLineStyle.xlContinuous
A.Weight = XlBorderWeight.xlMedium
End Sub
