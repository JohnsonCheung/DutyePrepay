Attribute VB_Name = "nVb_Vb"
Option Compare Database
Option Explicit
Public IsBch As Boolean
Public Const C_Lib$ = "DutyPrepay5"

Function AskQuit() As Boolean
If MsgBox("Quit?", vbYesNo + vbDefaultButton2) = vbYes Then Application.Quit
End Function

Function Cfn(Msg$) As Boolean
Stop
End Function

Function DigCnt%(A&)
DigCnt = Len(CStr(A))
End Function

Sub Done()
If IsBch Then Exit Sub
MsgBox "Done"
End Sub

Function Max(A, ParamArray Ap())
Dim Av(), O
Av = Ap
O = A
Dim J%
For J% = 0 To UB(Av)
    If O < Av(J) Then O = Av(J)
Next
Max = O
End Function

Function Min(A, ParamArray Ap())
Dim Av(), O
Av = Ap
O = A
Dim J&
For J = 0 To UB(Av)
    If O > Av(J) Then O = Av(J)
Next
Min = O
End Function

Function Pipe(ParamArray Ap())
Dim Av()
Av = Ap
Dim J%
Dim O
Dim R
Dim Mth$
VarAsg Av(0), O
For J = 1 To UB(Av)
    Mth = Av(J)
    VarAsg Run(Mth, O), R
    VarAsg R, O
Next
Pipe = O
End Function

Function RunAv(Fn$, Av())
Dim O
Select Case Sz(Av)
Case 0: O = Run(Fn)
Case 1: O = Run(Fn, Av(0))
Case 2: O = Run(Fn, Av(0), Av(1))
Case 3: O = Run(Fn, Av(0), Av(1), Av(2))
Case 4: O = Run(Fn, Av(0), Av(1), Av(2), Av(3))
Case 5: O = Run(Fn, Av(0), Av(1), Av(2), Av(3), Av(4))
Case 6: O = Run(Fn, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5))
Case 7: O = Run(Fn, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6))
Case 8: O = Run(Fn, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6), Av(7))
Case Else: Stop
End Select
RunAv = O
End Function

Function Start(Optional VBarMsg$ = "Start?", Optional Tit$ = "Start?") As Boolean
'If IsBch Then Start = True: Exit Function
Start = MsgBox(RplVBar(VBarMsg), vbQuestion + vbYesNo + vbDefaultButton1, Tit) = vbYes
End Function

Function TimStmp$(Optional Pfx$)
Static I&
Dim A$
If Pfx <> "" Then A = Pfx & "-"
TimStmp = A & Format(Now(), "YYYY-MM-DD_HHMMSS-") & I
I = I + 1
End Function

Function TryRun(Fct$, ParamArray Ap()) As OptV
Dim Av(): Av = Ap
Dim V
On Error GoTo X
V = RunAv(Fct, Av)
TryRun = OptVNew(V)
X:
End Function

Sub TryRun__Tst()
Dim V As OptV: V = TryRun("VarLng", "slfk")
Stop
End Sub

Function TryRunAv(Fn$, Av()) As OptV
On Error GoTo X
Dim O
Select Case Sz(Av)
Case 0: O = Run(Fn)
Case 1: O = Run(Fn, Av(0))
Case 2: O = Run(Fn, Av(0), Av(1))
Case 3: O = Run(Fn, Av(0), Av(1), Av(2))
Case 4: O = Run(Fn, Av(0), Av(1), Av(2), Av(3))
Case 5: O = Run(Fn, Av(0), Av(1), Av(2), Av(3), Av(4))
Case 6: O = Run(Fn, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5))
Case 7: O = Run(Fn, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6))
Case 8: O = Run(Fn, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6), Av(7))
Case Else: Stop
End Select
TryRunAv = OptVNew(O)
Exit Function
X:

End Function

Function Version$()
Version = "Verision 2007-03-14@0111"
End Function
