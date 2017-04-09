Attribute VB_Name = "nFs_TmpPth"
Option Compare Database
Option Explicit

Function TmpPth$(Optional SubFdr$ = "")
Dim A$
If SubFdr <> "" Then A = SubFdr & "\"
Dim O$
    O = Fso.GetSpecialFolder(TemporaryFolder) & "\" & A
If A <> "" Then PthEns O
TmpPth = O
End Function

Sub TmpPthClr()
PthClr TmpPth
End Sub
