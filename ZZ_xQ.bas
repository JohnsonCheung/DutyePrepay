Attribute VB_Name = "ZZ_xQ"

Option Compare Text
Option Explicit
Const cMod$ = cLib & ".xQ"

Function Q_Ln(oLn_wQuote$, pLn$, Optional pQ$ = CtSngQ) As Boolean
'Aim: Quote each name in {pLn} by {pQ} into {oLn_wQuote}
Const cSub$ = "Q_Ln"
If pLn$ = "" Then oLn_wQuote = "": Exit Function
Dim An$(): An = Split(pLn, CtComma)
Dim A$: A = Q_S(Rmv_Q(Trim(An(0)), pQ), pQ)
Dim J%: For J = 1 To Sz(An) - 1
    A = A & CtComma & Q_S(Trim(An(J)), pQ)
Next
oLn_wQuote = A$
Exit Function
E: Q_Ln = True
End Function

Function Q_SqBkt$(pS$)
If Left(pS, 1) = "(" And Right(pS, 1) = ")" Then Q_SqBkt = pS: Exit Function
Q_SqBkt = Q_S(pS, "[]")
End Function

Function Q_SqBkt__Tst()
Debug.Print Q_SqBkt("[dsfdf]")
Debug.Print Q_SqBkt("dsfdf")
End Function
