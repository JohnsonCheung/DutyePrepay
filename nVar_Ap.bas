Attribute VB_Name = "nVar_Ap"
Option Compare Database
Option Explicit

Function ApAv(ParamArray Ap()) As Variant()
ApAv = Ap
End Function

Function ApAy(ParamArray Ap()) As Variant()
Dim Av(): Av = Ap
Dim V, O()
For Each V In Av
    If IsArray(V) Then
        PushAy O, V
    Else
        Push O, V
    End If
Next
ApAy = O
End Function

Function ApIntAy(ParamArray Ap()) As Integer()
Dim Av(): Av = Ap
Dim U&: U = UB(Av)
Dim O%(): ReSz O, U
Dim J&
For J = 0 To U
    O(J) = Av(J)
Next
ApIntAy = O
End Function

Function ApJnComma$(ParamArray Ap())
Dim Av(): Av = Ap
ApJnComma = AyJn(Av, CtComma)
End Function

Function ApLngAy(ParamArray Ap()) As Long()
Dim Av(): Av = Ap
Dim U&: U = UB(Av)
Dim O&(): ReSz O, U
Dim J&
For J = 0 To U
    O(J) = Av(J)
Next
ApLngAy = O
End Function

Function ApNonBlank(ParamArray Ap())
Dim J%
For J = 0 To UBound(Ap)
    If Not IsMissing(Ap(J)) Then
        If Nz(Ap(J), "") <> "" Then
            ApNonBlank = Ap(J)
            Exit Function
        End If
    End If
Next
End Function

Function ApNonBlank___Tst()
Debug.Print TypeName(ApNonBlank("", 1))
End Function

Function ApSy(ParamArray Itm_or_Ay_Ap()) As String()
Dim Av(): Av = Itm_or_Ay_Ap
Dim OAy(): OAy = AyExpandAy(Av)
ApSy = AySy(OAy)
End Function
