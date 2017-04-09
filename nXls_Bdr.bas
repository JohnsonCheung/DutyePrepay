Attribute VB_Name = "nXls_Bdr"
Option Compare Database
Option Explicit

Sub BdrSetVLin(A As Border)
A.LineStyle = XlLineStyle.xlContinuous
A.Weight = XlBorderWeight.xlMedium
End Sub
