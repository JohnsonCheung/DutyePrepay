Attribute VB_Name = "nStr_Rpl"
Option Compare Database
Option Explicit

Function RplCr$(S)
RplCr = Replace(S, vbCr, " ")
End Function

Function RplCrLf$(S)
RplCrLf = RplCr(RplLf(S))
End Function

Function RplDblSpc$(S)
Dim O$: O = S
While InStr(O, "  ") > 0
    O = Replace(O, "  ", " ")
Wend
RplDblSpc = O
End Function

Function RplLf$(S)
RplLf = Replace(S, vbLf, " ")
End Function

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
