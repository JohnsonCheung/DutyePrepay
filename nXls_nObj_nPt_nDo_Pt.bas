Attribute VB_Name = "nXls_nObj_nPt_nDo_Pt"
Option Compare Database
Option Explicit

Sub PtSrt(Ppt As PivotTable)
Dim I As PivotField
For Each I In Ppt.PivotFields
    With I
        If .Name <> "Data" Then .AutoSort xlAscending, .Name
    End With
Next
End Sub


