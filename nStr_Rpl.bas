Attribute VB_Name = "nStr_Rpl"
Option Compare Database
Option Explicit

Function RplPun$(S)
Dim O$, J&, C$
For J = 1 To Len(S)
    C = Mid(S, J, 1)
    If ChrIsPun(C) Then
        O = O + " "
    Else
        O = O + C
    End If
Next
RplPun = O
End Function

Function RplVBar$(S)
RplVBar = Replace(S, "|", vbCrLf)
End Function
