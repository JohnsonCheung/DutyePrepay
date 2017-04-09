Attribute VB_Name = "nStr_Align"
Option Compare Database
Option Explicit

Function AlignL$(S, W%)
Dim L%: L = W - Len(S)
If L >= 0 Then
    AlignL = S & Space(L)
Else
    AlignL = S & " "
End If
End Function

Function AlignR$(S, W%)
Dim L%: L = W - Len(S)
If L >= 0 Then
    AlignR = Space(L) & S
Else
    AlignR = " " & S
End If
End Function
