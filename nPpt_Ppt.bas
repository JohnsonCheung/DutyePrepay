Attribute VB_Name = "nPpt_Ppt"
Option Compare Database
Option Explicit

Sub PptCls(A As PowerPoint.Presentation, Optional pSav As Boolean)
Dim P As PowerPoint.Application: Set P = A.Application
P.DisplayAlerts = False
A.Save
A.Close
P.DisplayAlerts = True
End Sub

