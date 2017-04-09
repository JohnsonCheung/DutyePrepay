Attribute VB_Name = "nStr_Parse"
Option Compare Database
Option Explicit

Function ParseNm$(A$)
Dim J%, C$
C = Left(A, 1)
If Not ChrIsLetter(C) Then Exit Function
For J = 2 To Len(A)
    C = Mid(A, J, 1)
    If Not ChrIsNmChr(C) Then
        ParseNm = Left(A, J - 1)
        A = Mid(A, J)
        Exit Function
    End If
Next
ParseNm = A
A = ""
End Function

Sub ParseNm__Tst()
Dim A$, Act$
A = "sdkf$lksdf"
Act = ParseNm(A)
Debug.Assert Act = "sdkf"
Debug.Assert A = "$lksdf"
End Sub

Function ParseStr$(A$, S)
If IsPfx(A, S) Then
    ParseStr = S
    A = RmvPfx(A, S)
End If
End Function

Function ParseSy$(A$, Ay$())
Dim I, O$
For Each I In Ay
    O = ParseStr(A, I)
    If O <> "" Then
        ParseSy = O
        Exit Function
    End If
Next
End Function

Function ParseTillClsBkt(A$)
Dim P%: P = InStr(A, ")"): If P = 0 Then Er "ParseTillClsBkt: No [)] in given {A}", A
Dim N%, J%, C$
For J = 1 To Len(A)
    C = Mid(A, J, 1)
    If C = "(" Then
        N = N + 1
    ElseIf C = ")" Then
        If N = 0 Then
            ParseTillClsBkt = Left(A, J - 1)
            A = Mid(A, J + 1)
            Exit Function
        End If
        N = N - 1
        If N < 0 Then Er "ParseTillClsBkt: Impossible"
    End If
Next
End Function
