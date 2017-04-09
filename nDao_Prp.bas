Attribute VB_Name = "nDao_Prp"
Option Compare Database
Option Explicit

Sub PrpDrp(PrpNm$, A As DAO.Properties)
If PrpIsExist(PrpNm, A) Then
    A.Delete PrpNm
End If
End Sub

Function PrpIsExist(PrpNm$, A As DAO.Properties) As Boolean
Dim I As Property
For Each I In A
    If I.Name = PrpNm Then PrpIsExist = True: Exit Function
Next
End Function
