Attribute VB_Name = "nXls_nObj_nLoNm_LoNm"
Option Compare Database
Option Explicit

Function LoNmNz$(LoNm$, A As Worksheet)
If LoNm <> "" Then LoNmNz = LoNm: Exit Function
Dim J%, N$
For J = 1 To 1000
    N = "Tbl_" & J
    If Not WsHasLoNm(A, N) Then LoNm = N: Exit Function
Next
Er "LoNmNz:Impossible"
End Function
