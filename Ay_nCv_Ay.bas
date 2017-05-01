Attribute VB_Name = "Ay_nCv_Ay"
Option Compare Database
Option Explicit

Function AyIntAy(Ay) As Integer()
If AyIsEmpty(Ay) Then Exit Function
If VarType(Ay(0)) = vbInteger Then AyIntAy = Ay: Exit Function
AyIntAy = AyInto(Ay, EmptyIntAy)
End Function

Function AyInto(Ay, OInto)
Dim U&: U = UB(Ay)
Dim O$(): ReSz O, U
Dim J&
For J = 0 To U
    O(J) = Ay(J)
Next
AyInto = O
End Function

Function AySy(Ay) As String()
Dim U&: U = UB(Ay)
Dim O$(): ReSz O, U
Dim J&
For J = 0 To U
    O(J) = VarToStr(Ay(J))
Next
AySy = O
End Function
