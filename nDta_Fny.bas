Attribute VB_Name = "nDta_Fny"
Option Compare Database
Option Explicit

Function FnyHtm$(Fny$())
Dim O$, J%
O = "<th>"
Dim F$(): F = AyMapIntoSy(Fny, "CamelNrm")
For J = 0 To UB(F)
    O = O & "<td>" & F(J)
Next
FnyHtm = O
End Function

Function FnySel(Fny$(), StarFnStr$, Optional IsChkStarFnStrMustGood As Boolean) As String()
Dim O$()
    O = NmstrBrk(StarFnStr)
    
'--- Check Star must be at end
    Dim J%, IsEr As Boolean
    For J = 0 To UB(Fny) - 1
        If Fny(J) = "*" Then IsEr = True: Exit For
    Next
    If IsEr Then Er "FnySel: [*] in {StarFnStr} be last one", StarFnStr

'--- Check only one
If IsChkStarFnStrMustGood Then
    Dim X$()
    X = AyRmvEle(AyMinus(O, Fny), "*")
    If Not AyIsEmpty(X) Then
        Er "FnySel: {StarFnStr} has {Fields} not in {Fny}", StarFnStr, FnyToStr(X), FnyToStr(Fny)
    End If
End If

If LasEle(O) = "*" Then
    Pop O
    PushAy O, AyMinus(Fny, O)
End If
FnySel = O
End Function

Function FnyToStr$(Fny$())
Dim J%, O$()
O = Fny
For J = 0 To UB(Fny)
    If Not StrIsNm(Fny(J)) Then O(J) = Quote(Fny(J), "[]")
Next
FnyToStr = Join(O, " ")
End Function

Private Sub FnySel__Tst()
Dim A1$(): A1 = LvsSplit("A B C D E")
Dim A2$: A2 = "E D A *"
Dim F$()
F = FnySel(A1, A2)
Debug.Assert Sz(F) = 5
Debug.Assert F(0) = "E"
Debug.Assert F(1) = "D"
Debug.Assert F(2) = "A"
Debug.Assert F(3) = "B"
Debug.Assert F(4) = "C"
'=====
A2 = "G *"
F = FnySel(A1, A2)
Debug.Assert Sz(F) = 6
Debug.Assert F(0) = "G"
Debug.Assert F(1) = "A"
Debug.Assert F(2) = "B"
Debug.Assert F(3) = "C"
Debug.Assert F(4) = "D"
Debug.Assert F(5) = "E"
'=====

A2 = "G *"
GoSub ShouldThrowEr
Exit Sub

ShouldThrowEr:
    On Error GoTo Er1
    F = FnySel(A1, A2, IsChkStarFnStrMustGood:=True)
    Debug.Assert False
Er1:
    Return

End Sub
