Attribute VB_Name = "nXls_nObj_nPt_nDo_Pt"
Option Compare Database
Option Explicit

Sub PtSrt(pPt As PivotTable)
Dim I As PivotField
For Each I In pPt.PivotFields
    With I
        If .Name <> "Data" Then .AutoSort xlAscending, .Name
    End With
Next
End Sub


