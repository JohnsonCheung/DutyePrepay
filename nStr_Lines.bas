Attribute VB_Name = "nStr_Lines"
Option Compare Database
Option Explicit

Function LinesFst$(Lines$)
LinesFst = TakBef(Lines, vbLf)
End Function

Function LinesSplit(Lines$) As String()
LinesSplit = Split(Lines, vbCrLf)
End Function

Function LinesTrim$(S)
LinesTrim = LinesTrimEnd(LinesTrimBeg(S))
End Function

Function LinesTrimBeg$(S)
Dim O$: O = S
Dim C$, J&
For J = 1 To Len(O)
    C = Left(O, 1)
    If C = vbCr Then O = Mid(O, 2): GoTo Nxt
    If C = vbLf Then O = Mid(O, 2): GoTo Nxt
    Exit For
Nxt:
Next
LinesTrimBeg = O
End Function

Function LinesTrimEnd$(S)
Dim O$: O = S
Dim C$, J&
For J = Len(O) To 1 Step -1
    C = Right(O, 1)
    If C = vbCr Then O = Left(O, Len(O) - 1): GoTo Nxt
    If C = vbLf Then O = Left(O, Len(O) - 1): GoTo Nxt
    Exit For
Nxt:
Next
LinesTrimEnd = O
End Function
