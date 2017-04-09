Attribute VB_Name = "nDta_Ay"
Option Compare Database
Option Explicit

Function AySqH(Ay)
Dim O(), C&
C = Sz(Ay)
ReDim O(1 To 1, 1 To C)
Dim J&
For J = 0 To C - 1
    O(1, J + 1) = Ay(J)
Next
AySqH = O
End Function

Function AySqV(Ay)
Dim O(), R&
R = Sz(Ay)
ReDim O(1 To R, 1 To 1)
Dim J&
For J = 0 To R - 1
    O(J + 1, 1) = Ay(J)
Next
AySqV = O
End Function
