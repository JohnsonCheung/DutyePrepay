Attribute VB_Name = "nStr_Tak"
Option Compare Database
Option Explicit

Function TakAft$(S, Aft, Optional InclAft As Boolean)
Dim P&: P = InStr(S, Aft): If P = 0 Then Exit Function
If InclAft Then
    TakAft = Mid(S, P)
Else
    TakAft = Mid(S, P + Len(Aft))
End If
End Function

Function TakBef$(S, Bef, Optional InclBef As Boolean)
Dim P&: P = InStr(S, Bef): If P = 0 Then Exit Function
If InclBef Then
    TakBef = Left(S, P - 1 + Len(Bef))
Else
    TakBef = Left(S, P - 1)
End If
End Function
