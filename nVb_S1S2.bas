Attribute VB_Name = "nVb_S1S2"
Option Compare Database
Option Explicit
Type S1S2
    S1 As String
    S2 As String
End Type

Function S1S2New(S1, S2) As S1S2
S1S2New.S1 = S1
S1S2New.S2 = S2
End Function

Function S1S2Trim(P As S1S2) As S1S2
With P
    .S1 = Trim(.S1)
    .S2 = Trim(.S2)
End With
S1S2Trim = P
End Function
