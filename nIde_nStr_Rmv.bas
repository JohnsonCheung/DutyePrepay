Attribute VB_Name = "nIde_nStr_Rmv"
Option Compare Database
Option Explicit

Function RmvStrTok$(S)
Dim J%, IsInQ As Boolean
Dim O$, C$
For J = 1 To Len(S)
    C = Mid(S, J, 1)
    If C = """" Then
        IsInQ = Not IsInQ
    Else
        If Not IsInQ Then
            O = O + C
        End If
    End If
Next
RmvStrTok = O
End Function


