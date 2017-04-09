Attribute VB_Name = "nXls_nDta_Wb"
Option Compare Database
Option Explicit

Function WbAddDt(A As Workbook, Dt As Dt, Tn$) As Worksheet
Dim O As Worksheet
    Set O = WbAddWsAtEnd(A, Tn)
    DtPutCell Dt, WsA1(O)
Set WbAddDt = O
End Function
