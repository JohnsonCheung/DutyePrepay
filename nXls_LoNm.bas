Attribute VB_Name = "nXls_LoNm"
Option Compare Database
Option Explicit

Function LoNmIsExist(LoNm$, A As Worksheet) As Boolean
On Error GoTo X
LoNmIsExist = A.ListObjects(LoNm).Name = LoNm
X:
End Function

Function LoNmNz$(LoNm$, A As Worksheet)
If LoNm <> "" Then LoNmNz = LoNm: Exit Function
Dim J%, N$
For J = 1 To 1000
    N = "Tbl_" & J
    If Not LoNmIsExist(N, A) Then LoNm = N: Exit Function
Next
Er "LoNmNz:Impossible"
End Function
