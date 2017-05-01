Attribute VB_Name = "nVb_nNmstr_Nmstr"
Option Compare Database
Option Explicit

Function FnStrLvc$(FnStr$)
Dim F$(): F = NmBrk(FnStr)
Dim J%
For J = 0 To UB(F)
    If StrIsNm(F(J)) Then F(J) = Quote(F(J), "[]")
Next
FnStrLvc = Join(F, ",")
End Function

Function FnStrPkDic(FnStr) As Dictionary
Set FnStrPkDic = AyPkDic(NmBrk(FnStr))
End Function

Function NmBrk(NmStr) As String()
Dim M$, O$()
M = NmStr
With Brk1(M, "[")
    O = AyAdd(LvsSplit(.S1), O)
    If .S2 = "" Then NmBrk = O: Exit Function
    With Brk(.S2, "]")
        Push O, .S1
        PushAy O, NmBrk(.S2)
        NmBrk = O
        Exit Function
    End With
End With
NmBrk = O
End Function

Sub NmBrk__Tst()
Dim Act$(): Act = NmBrk(" skldf dfk   kdf [df d] kdf df [  a ]  ")
Dim Exp$(): Exp = ApSy("skldf", "dfk", "kdf", "df d", "kdf", "df", "a")
AyChkEq Act, Exp
End Sub

Function NmstrExcp(NmStr, Ay$()) As String()
NmstrExcp = AyMinus(Ay, NmstrExpd(NmStr, Ay))
End Function

Function NmstrExpd(NmStr, Ay$()) As String()
Dim Ny$(): Ny = NmBrk(NmStr)
If AyIsEmpty(Ny) Then Exit Function
Dim I, O$()
For Each I In Ny
    PushAyNoDup O, NmExpd(I, Ay)
Next
NmstrExpd = O
End Function
