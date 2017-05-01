Attribute VB_Name = "nAy_nToStr"
Option Compare Database
Option Explicit

Function BoolAyToStr$(A() As Boolean)
If AyIsEmpty(A) Then Exit Function
Dim O$(), J&, I
For Each I In A
    O(J) = IIf(I, 1, 0)
    J = J + 1
Next
BoolAyToStr = Jn(O, " ")
End Function

Function BytAyToStr$(A() As Byte)
BytAyToStr = Jn(A, " ")
End Function

Function LngAyToStr$(A&())
LngAyToStr = Jn(A, " ")
End Function
