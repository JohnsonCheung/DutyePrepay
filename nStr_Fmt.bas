Attribute VB_Name = "nStr_Fmt"
Option Compare Database
Option Explicit

Function Fmt$(FmtStr$, ParamArray Ap())
Dim O$, I%, V
O = Replace(FmtStr, "|", vbCrLf)
I = 0
For Each V In Ap
    O = Replace(O, "{" & I & "}", Nz(V, "Null")): I = I + 1
Next
Fmt = O
End Function

Function Fmt_yMmmWww(Dte As Date) As String
If Dte < Date Then
    Fmt_yMmmWww = " Past"
    Exit Function
End If
Fmt_yMmmWww = Right(Year(Dte), 1) & "M" & Format(Month(Dte), "00") & "W" & Format(DteWkNo(Dte), "00")
End Function

Function FmtDic$(NmMacro$, Dic As Dictionary)
'Aim: pFmtStr is in forAt of xxxx{Fld1}xxx{Fld2}.  Return the subst string by subst the fields in {pRs} into {pFmtStr}
Dim S$: S = NmMacro
Dim K
Dim A$
For Each K In Dic
    A = "{" & K & "}"
    S = Replace(S, A, Dic(K))
Next
FmtDic = S
End Function

Function FmtNm$(NmStr$, ParamArray Ap())
Dim Av(): Av = Ap
FmtNm = FmtNmAv(NmStr, Av)
End Function

Sub FmtNm__Tst()
Dim D As Date
D = #1/1/2016 1:23:44 PM#
Debug.Assert FmtNm("dsf{0},{1}dlf", 1, D) = "dsf1,2016-01-01 13:23:44dlf"
End Sub

Function FmtNmAv$(MacroStr, Av())
Dim O$: O = MacroStr
Dim A$(): A = BrkMacroStr(O)
Dim J&
For J = 0 To UB(A)
    O = Replace(O, A(J), Av(J))
Next
FmtNmAv = O
End Function

Sub FmtNmAv__Tst()
Debug.Assert FmtNm("{a}--{b}...{a}!!!", 1, 2) = "1--2...1!!!"
End Sub

Function FmtQQ$(QQStr$, ParamArray Ap())
Dim Av(): Av = Ap
FmtQQ = FmtQQAv(QQStr, Av)
End Function

Function FmtQQAv(QQStr$, Av())
Dim J%
Dim S$: S = QQStr
For J = 0 To UBound(Av)
    S = Replace(S, "?", Av(J), 1, 1)
Next
FmtQQAv = S
End Function

Function FmtQQVBar$(QQStr$, ParamArray Ap())
Dim Av(): Av = Ap
FmtQQVBar = RplVBar(FmtQQAv(QQStr, Av))
End Function
