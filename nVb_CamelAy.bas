Attribute VB_Name = "nVb_CamelAy"
Option Compare Database
Option Explicit

Function CamelAy(Camel) As String()
'--- A*
Dim AAy$()
    Dim A$: A = Camel
    Dim J%, Seg$, C$
    Seg = FstChr(A)
    For J = 2 To Len(A)
        C = Mid(A, J, 1)
        If ChrIsCap(C) Then
            Push AAy, Seg
            Seg = C
        Else
            Seg = Seg + C
        End If
    Next
    Push AAy, Seg
'
Dim O$()
    Dim LasIdx%
    Push O, AAy(0)
    For J = 1 To UB(AAy)
        If Len(AAy(J)) = 1 Then
            LasIdx = UB(O)
            O(LasIdx) = O(LasIdx) & AAy(J)
        Else
            Push O, AAy(J)
        End If
    Next
CamelAy = O
End Function

Sub CamelAy__Tst()
'1 Declare
Dim Camel
Dim Act$()
Dim Exp$()

'2 Assign
Camel = "AyBrk"
Exp = ApSy("Ay", "Brk")

'3 Calling
Act = CamelAy(Camel)

'4 Asst
AyAsstEq Act, Exp


'2 Assign
Camel = "GL"
Exp = ApSy("GL")

'3 Calling
Act = CamelAy(Camel)

'4 Asst
AyAsstEq Act, Exp
End Sub

Function CamelNrm$(Camel)
CamelNrm = Join(CamelAy(Camel), " ")
End Function

Sub CamelNrm__Tst()
Debug.Assert CamelNrm("GL") = "GL"
Debug.Assert CamelNrm("GLA") = "GLA"
Debug.Assert CamelNrm("GLAy") = "GL Ay"
End Sub
