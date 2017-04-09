Attribute VB_Name = "nEr_Er"
Option Compare Database
Option Explicit

Sub Er(Msg$, ParamArray Ap())
Dim Av(): Av = Ap
Dim A As Dt: A = ErNewMsgAv(Msg, Av)
DtBrw A
Err.Raise 1
End Sub

Function ErApd(A As Dt, Msg$, ParamArray Ap()) As Dt
Dim ErDr(): ErDr = ApAy(Msg, Ap)
ErApd = ErApdDr(A, ErDr)
End Function

Sub ErApd__Tst()
Dim A As Dt:
A = ErNew("sldfjsdf", 1, 2, 3, Now, "sdf")
A = ErApd(A, "skldfjsdf")
DtBrw A
End Sub

Function ErApdDr(A As Dt, ErDr()) As Dt
ErReSzFny A, UB(ErDr)
ErApdDr = DtApdDr(A, ErDr)
End Function

Function ErApdEr(A As Dt, Er As Dt) As Dt
If ErIsSom(A) Then
    ErApdEr = DtUnion(A, Er)
Else
    ErApdEr = DtUnion(DtNew(ApSy("Msg"), Array()), Er)
End If
End Function

Sub ErAsst(A As Dt, Optional ExplainMsgAv)
If DtNRec(A) = 0 Then Exit Sub
Dim E As Dt: E = ErNewDr(ExplainMsgAv)
DtBrw ErApdEr(E, A)
Err.Raise 1
End Sub

Function ErExplain(A As Dt, Msg$, ParamArray Ap()) As Dt
If ErIsNone(A) Then Exit Function
Dim Er As Dt
Dim Av(): Av = Ap
Er = ErNewMsgAv(Msg, Av)
ErExplain = ErApdEr(Er, A)
End Function

Function ErIsNone(A As Dt) As Boolean
ErIsNone = Not ErIsSom(A)
End Function

Function ErIsSom(ErDt As Dt) As Boolean
ErIsSom = DtNRec(ErDt) > 0
End Function

Function ErNew(Msg$, ParamArray Ap()) As Dt
Dim Av(): Av = Ap
ErNew = ErNewMsgAv(Msg, Av)
End Function

Sub ErNew__Tst()
Dim Av()
DtBrw ErNew("AAAA")
'DtBrw ErNew("sldfjsdf", 1, 2, 3, Now, "sdf")
End Sub

Function ErNewDr(Dr) As Dt
If VarIsBlank(Dr) Then Exit Function
Dim OFny$()
Push OFny, "Msg"
Dim J%
For J = 0 To UB(Dr) - 1
    Push OFny, "V" & J
Next
Dim ODrAy()
Push ODrAy, Dr
ErNewDr = DtNew(OFny, ODrAy, "Er")
End Function

Function ErNewDrAy(ErDrAy()) As Dt
Dim U&: U = UB(ErDrAy)
Dim Fny$(): ReDim Fny(U)
Dim J&
Fny(0) = "Msg"
For J = 1 To U
    Fny(J) = "v" & J - 1
Next
ErNewDrAy = DtNew(Fny, ErDrAy)
End Function

Function ErNewMsgAv(Msg$, Av()) As Dt
'Aim: The return Dt is [Msg V0 V1 ..]
Dim U%: U = UB(Av)
Dim F$()
    Dim J%
    Push F, "Msg"
    For J = 0 To U
        Push F, "V" & J
    Next
Dim Dr()
    Push Dr, Msg
    PushAy Dr, Av
ErNewMsgAv = DtNew(F, Array(Dr))
End Function

Sub ErNewMsgAv__Tst()
Dim Av()
DtBrw ErNewMsgAv("aaaa", Av)
End Sub

Sub ErReSzFny(A As Dt, NewFnyUB&)
Dim UFny&: UFny = UB(A.Fny)
If NewFnyUB <= UFny Then Exit Sub
ReSz A.Fny, NewFnyUB
Dim J&
For J = UFny + 1 To NewFnyUB
    A.Fny(J) = "V" & J - 1
Next
End Sub
