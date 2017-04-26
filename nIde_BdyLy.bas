Attribute VB_Name = "nIde_BdyLy"
Option Compare Database
Option Explicit

Function _
    BdyLyLin$(BdyLy$(), Idx%)
Dim O$
Dim L$
Dim J%
Dim IsInContinue
For J = Idx To UB(BdyLy)
    L = BdyLy(J)
    If LasChr(L) = "_" Then
        IsInContinue = True
        O = O & LTrim(RmvLasChr(L))
    Else
        If IsInContinue Then
            BdyLyLin = O & LTrim(L)
        Else
            BdyLyLin = O & L
        End If
        Exit Function
    End If
Next
BdyLyLin = O
End Function

Function BdyLyMthLy(BdyLy$()) As String()
Dim J%, L$, O$()
For J = 0 To UB(BdyLy)
    L = BdyLyLin(BdyLy, J)
    If SrcLinIsMth(L) Then Push O, L
Next
BdyLyMthLy = O
End Function

