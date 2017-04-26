Attribute VB_Name = "nStr_Parse"
Option Compare Database
Option Explicit

Function ParseChr$(A$, ChrStr$)
Dim F$: F = FstChr(A)
Dim J%
For J = 1 To Len(ChrStr)
    If F = Mid(ChrStr, J, 1) Then ParseChr = Mid(ChrStr, J, 1): Exit Function
Next
End Function

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

Function ParseStr(A$, S) As Boolean
Dim O As Boolean
O = IsPfx(A, S)
If O Then A = RmvPfx(A, S)
ParseStr = O
End Function

Function ParseSy$(A$, Ay$())
Dim I, O$
For Each I In Ay
    If ParseStr(A, I) Then ParseSy = I: Exit Function
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
