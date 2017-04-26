Attribute VB_Name = "nEr_Er"
Option Compare Database
Option Explicit

Sub Er(Msg$, ParamArray Ap())
Dim Av(): Av = Ap
Dim A(): A = ErNewMsgAv(Msg, Av)
ErBrw A
Err.Raise 1
End Sub

Function ErApd(Er(), Msg$, ParamArray Ap()) As Variant()
Dim Av(): Av = Ap
ErApd = AyAdd(Er, ErNewMsgAv(Msg, Av))
End Function

Sub ErApd__Tst()
Dim A()
A = ErNew("sldfjsdf", 1, 2, 3, Now, "sdf")
A = ErApd(A, "skldfjsdf", 2)
ErBrw A
End Sub

Sub ErAsst(Er(), Optional MsgAv)
If AyIsEmpty(Er) Then Exit Sub
ErBrw AyAddItm(Er, MsgAv)
Err.Raise 1
End Sub

Sub ErBrw(Er)
'[Chk] is a table with or without records.  If no record, it means nothing to check.
'If there is record, something needs to be checked.
'So ChkBrw is written as this.
If AyIsEmpty(Er) Then Exit Sub
DtBrw ErDt(Er), "Please-check"
Err.Raise 1
End Sub

Function ErDt(Er) As Dt
Dim U%: U = DrAyColUB(Er)
Dim Fny$()
Push Fny, "Msg"
Dim J%
For J = 1 To U
    Push Fny, "v" & J - 1
Next
ErDt = DtNew(Fny, Er)
End Function

Function ErNew(Msg$, ParamArray Ap()) As Variant()
Dim Av(): Av = Ap
ErNew = ErNewMsgAv(Msg, Av)
End Function

Sub ErNew__Tst()
'ErBrw ErNew("AAAA")
ErBrw ErNew("sldfjsdf", 1, 2, 3, Now, "sdf")
End Sub

Function ErNewMsgAv(Msg$, Av) As Variant()
'Aim: The return Dt is [Msg V0 V1 ..]
Dim U%: U = UB(Av)
Dim Dr()
    Push Dr, Msg
    PushAy Dr, Av
ErNewMsgAv = Array(Dr)
End Function

Sub ErNewMsgAv__Tst()
Dim Av()
ErBrw ErNewMsgAv("sdfsdf", Array(1, 2, 3))
End Sub
