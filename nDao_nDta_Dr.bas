Attribute VB_Name = "nDao_nDta_Dr"
Option Compare Database
Option Explicit

Function DrHtm$(Dr)
Dim O$(), V
Push O, "<tr>"
If Not AyIsEmpty(Dr) Then
    For Each V In Dr
        Push O, "<td>" & VarToStr(V)
    Next
End If
DrHtm = Join(O, "")
End Function

Sub DrUpdRs(Dr, Rs As Recordset, RsIdx&())
Dim J%
For J = 0 To UB(RsIdx)
    Rs(RsIdx(J)) = Dr(J)
Next
End Sub
