Attribute VB_Name = "nVb_Re"
Option Compare Database
Option Explicit
Private X_MdSrchRe As RegExp

Function ReMch(Re As RegExp, Src) As MatchCollection
Set ReMch = Re.Execute(Src)
End Function

Sub ReMch__Tst()
'1 Declare
Dim Re As RegExp
Dim Src
Dim Act As MatchCollection
Dim Exp As MatchCollection

'2 Assign
Set Re = ReNew("^aa")
Src = "aaabb"

'3 Calling
Set Act = ReMch(Re, Src)

'4 Asst
Debug.Assert Act.Count = 1
Dim Act0 As IMatch2
Set Act0 = Act(0)
Debug.Assert Act0.Length = 2
Debug.Assert Act0.Value = "aa"
Debug.Assert Act0.FirstIndex = 0
Debug.Assert Act0.SubMatches.Count = 0
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
'ReRpl = X.Replace(Src, Rpl)
End Function

Function ReTest(S, Pattern$, Optional Glob As Boolean, Optional IgnoreCase As Boolean, Optional MulLin As Boolean) As Boolean
ReSetup Pattern, Glob, IgnoreCase, MulLin
'ReTest = X.Test(S)
End Function

Private Sub ReSetup(Pattern$, Optional Glob As Boolean, Optional IgnoreCase As Boolean, Optional MulLin As Boolean)
'With X
'    .Pattern = Pattern
'    .Global = Glob
'    .IgnoreCase = IgnoreCase
'    .MultiLine = MulLin
'End With
End Sub
