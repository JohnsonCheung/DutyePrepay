Attribute VB_Name = "nStr_Lvs"
Option Compare Database
Option Explicit

Function LvsSplit(Lvs) As String()
Dim O$: O = Trim(Lvs)
Dim P%:
Do
    P = InStr(O, "  ")
    If P = 0 Then LvsSplit = Split(O): Exit Function
    O = Replace(O, "  ", " ")
Loop
End Function
