Attribute VB_Name = "nVb_Re"
Option Compare Database
Option Explicit
Private X As New VBScript_RegExp_55.RegExp

Function ReMch(Src, Pattern$, Optional Glob As Boolean, Optional IgnoreCase As Boolean, Optional MulLin As Boolean) As MatchCollection
ReSetup Pattern, Glob, IgnoreCase, MulLin
Set ReMch = X.Execute(Src)
End Function

Sub ReMch__Tst()
'1 Declare
Dim Src
Dim Pattern$
Dim Glob As Boolean
Dim IgnoreCase As Boolean
Dim MulLin As Boolean
Dim Act As MatchCollection

'2 Assign
Src = "abcdef"
Pattern = "^abc"
Glob = False
IgnoreCase = True
MulLin = False

'3 Calling
Set Act = ReMch(Src, Pattern, Glob, IgnoreCase, MulLin)

'4 Asst
Debug.Assert Act.Count = 1
Dim Act0 As Match: Set Act0 = Act(0)
Debug.Assert Act0.FirstIndex = 0
Debug.Assert Act0.Value = "abc"
Debug.Assert Act0.Length = 3
End Sub

Function ReNew(Pattern$, Optional Glob As Boolean, Optional IgnoreCase As Boolean, Optional MulLin As Boolean) As RegExp
Dim O As New RegExp
With O
    .Pattern = Pattern
    .Global = Glob
    .IgnoreCase = IgnoreCase
    .MultiLine = MulLin
End With
Set ReNew = O
End Function

Function ReRpl$(Src, Rpl, Pattern$, Optional Glob As Boolean, Optional IgnoreCase As Boolean, Optional MulLin As Boolean)
ReSetup Pattern, Glob, IgnoreCase, MulLin
ReRpl = X.Replace(Src, Rpl)
End Function

Function ReTest(S, Pattern$, Optional Glob As Boolean, Optional IgnoreCase As Boolean, Optional MulLin As Boolean) As Boolean
ReSetup Pattern, Glob, IgnoreCase, MulLin
ReTest = X.Test(S)
End Function

Private Sub ReSetup(Pattern$, Optional Glob As Boolean, Optional IgnoreCase As Boolean, Optional MulLin As Boolean)
With X
    .Pattern = Pattern
    .Global = Glob
    .IgnoreCase = IgnoreCase
    .MultiLine = MulLin
End With
End Sub
