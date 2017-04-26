Attribute VB_Name = "nVb_nNmstr_Nmstr"
Option Compare Database
Option Explicit

Function FnStrLvc$(FnStr$)
Dim F$(): F = NmstrBrk(FnStr)
Dim J%
For J = 0 To UB(F)
    If StrIsNm(F(J)) Then F(J) = Quote(F(J), "[]")
Next
FnStrLvc = Join(F, ",")
End Function

Function FnStrPkDic(FnStr) As Dictionary
Set FnStrPkDic = AyPkDic(NmstrBrk(FnStr))
End Function

Function NmstrBrk(Nmstr) As String()
Dim M$, O$()
M = Nmstr
With StrBrk1(M, "[")
    O = AyAdd(LvsSplit(.S1), O)
    If .S2 = "" Then NmstrBrk = O: Exit Function
    With StrBrk(.S2, "]")
        Push O, .S1
        PushAy O, NmstrBrk(.S2)
        NmstrBrk = O
        Exit Function
    End With
End With
NmstrBrk = O
End Function

Sub NmstrBrk__Tst()
Dim Act$(): Act = NmstrBrk(" skldf dfk   kdf [df d] kdf df [  a ]  ")
Dim Exp$(): Exp = ApSy("skldf", "dfk", "kdf", "df d", "kdf", "df", "a")
AyChkEq Act, Exp
End Sub

Function NmstrExcp(Nmstr, Ay$()) As String()
NmstrExcp = AyMinus(Ay, NmstrExpd(Nmstr, Ay))
End Function

Function NmstrExpd(Nmstr, Ay$()) As String()
Dim Ny$(): Ny = NmstrBrk(Nmstr)
If AyIsEmpty(Ny) Then Exit Function
Dim I, O$()
For Each I In Ny
    PushAyNoDup O, NmExpd(I, Ay)
Next
NmstrExpd = O
End Function
