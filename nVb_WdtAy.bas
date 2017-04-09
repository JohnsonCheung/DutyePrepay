Attribute VB_Name = "nVb_WdtAy"
Option Compare Database
Option Explicit

Function WdtAyHdr$(WdtAy%(), Fny$())
Dim U1%: U1 = UB(Fny): If U1 = -1 Then Exit Function
Dim U2%: U2 = UB(WdtAy)
Dim U%:   U = Max(U1, U2)
Dim O$(): ReDim O(U)
Dim J%
For J = 0 To U
    If J > U1 Then
        O(J) = Space(WdtAy(J))
    Else
        O(J) = AlignL(Fny(J), WdtAy(J))
    End If
Next
WdtAyHdr = Quote(Join(O, " | "), "| * |")
End Function

Function WdtAyLin$(WdtAy%())
Dim U%: U = UB(WdtAy)
If U = -1 Then Exit Function
Dim O$(): ReDim O(U)
Dim J%
For J = 0 To U
    O(J) = String(WdtAy(J) + 2, "-")
Next
WdtAyLin = Quote(Join(O, "|"), "|")
End Function
