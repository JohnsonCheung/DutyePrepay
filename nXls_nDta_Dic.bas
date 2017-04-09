Attribute VB_Name = "nXls_nDta_Dic"
Option Compare Database
Option Explicit

Sub DicPutCell(A As Dictionary, Cell As Range)
SqPutCell DicSq(A), Cell
End Sub

Function DicSq(A As Dictionary) As Variant()
If DicIsEmpty(A) Then Exit Function
Dim N&: N = A.Count
Dim O(): ReDim O(1 To N, 1 To 2)
Dim R&: R = 0
Dim K
For Each K In A.Keys
    R = R + 1
    O(R, 1) = K
    O(R, 2) = A(K)
Next
DicSq = O
End Function

