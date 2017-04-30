Attribute VB_Name = "nXls_nObj_nPt_nInf_Pt"
Option Compare Database
Option Explicit
Function PtNmNz$(PtNm$, A As Worksheet)
If PtNm <> "" Then PtNmNz = PtNm: Exit Function
Dim J%
For J = 1 To 1000
    If Not PtNmIsExist("PT_" & J, A) Then PtNmNz = "PT_" & J: Exit Function
Next
Er "PtNmNz: Impossible"
End Function
