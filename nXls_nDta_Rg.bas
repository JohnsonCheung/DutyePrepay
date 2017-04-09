Attribute VB_Name = "nXls_nDta_Rg"
Option Compare Database
Option Explicit

Function RgDt(A As Range, Optional Tn$ = "Table") As Dt
Dim Sq: Sq = A.Value
Dim Fny$()
    Fny = AySy(SqDr(Sq, 1))
Dim DrAy()
    Dim J&
    For J = 2 To UBound(Sq, 1)
        Push DrAy, SqDr(Sq, J)
    Next
RgDt = DtNew(Fny, DrAy, Tn)
End Function
