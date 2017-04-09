Attribute VB_Name = "nIde_nTok_Md"
Option Compare Database
Option Explicit

Function LinRmvRmk$(Lin)
If InStr(Lin, "'") = 0 Then LinRmvRmk = Lin: Exit Function
Dim J%, IsInQ As Boolean
Dim C$
For J = 1 To Len(Lin)
    C = Mid(Lin, J, 1)
    If C = "'" Then
        If Not IsInQ Then
            LinRmvRmk = Left(Lin, J - 1)
            Exit Function
        End If
    End If
    If C = """" Then
        IsInQ = Not IsInQ
    End If
Next
Er "LinRmvRmk: Impossible"
End Function

Function MdTokAy(Optional A As CodeModule) As String()
Dim Ly$(): Ly = MdLy(A)
Dim O$(), J%
For J = 0 To UB(Ly)
    PushAyNoDup O, LinTokAy(Ly(J))
Next
MdTokAy = O
End Function

Sub MdTokAy__Tst()
AyBrw AySrt(MdTokAy(Md("nIde_Md"))), WithIdx:=True
End Sub

Private Sub RmvStrTok__Tst()
Debug.Assert RmvStrTok("abc ""lksdjf"" 1234") = "abc  1234"
End Sub
