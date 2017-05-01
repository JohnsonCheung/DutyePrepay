Attribute VB_Name = "nDao_Prps"
Option Compare Database
Option Explicit

Function PrpsToStr$(A As DAO.Properties)
Dim mA$, I, P As Property, O$()
On Error GoTo R
For Each I In A
    Set P = I
    Push O, PrpToStr(P)
Next
PrpsToStr = Join(O, vbCrLf)
Exit Function
R: PrpsToStr = ErStr("PrpsToStr")
End Function

Sub PrpsToStr__Tst()
Dim Prps As DAO.Properties: Set Prps = Tbl("Permit").Properties
Debug.Print PrpsToStr(Prps)
End Sub
