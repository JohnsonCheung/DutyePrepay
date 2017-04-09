Attribute VB_Name = "nStr_Quote"
Option Compare Database
Option Explicit

Function Quote$(S, Q$)
Dim Q1$, Q2$
With QuoteBrk(Q)
    Q1 = .S1
    Q2 = .S2
End With
Quote = Q1 & S & Q2
End Function

Function QuoteBkt$(S)
QuoteBkt = Quote(S, "()")
End Function

Function QuoteBrk(Q$) As S1S2
Select Case Len(Q)
Case 0
Case 1: QuoteBrk = S1S2New(Q, Q)
Case 2: QuoteBrk = S1S2New(Left(Q, 1), Right(Q, 1))
Case Else:
    Dim P%: P = InStr(Q, "*")
    If P = 0 Then Er "{QuoteStr} is not a valid", Q
    QuoteBrk = S1S2New(Left(Q, P - 1), Mid(Q, P + 1))
End Select
End Function

Function QuoteRmv$(S, Optional Quote$ = "'")
Dim O$
With QuoteBrk(Quote)
    O = RmvPfx(S, .S1)
    O = RmvSfx(S, .S2)
End With
QuoteRmv = O
End Function

Function QuoteRmv__Tst()
Debug.Assert QuoteRmv("[aaaa]", "[]") = "aaaa"
End Function
