Attribute VB_Name = "nDao_nFld_FnStr"
Option Compare Database
Option Explicit

Function FnStrBrk(FnStr) As String()
Dim M$, O$()
M = FnStr
With StrBrk1(M, "[")
    O = AyAdd(LvsSplit(.S1), O)
    If .S2 = "" Then FnStrBrk = O: Exit Function
    With StrBrk(.S2, "]")
        Push O, .S1
        PushAy O, FnStrBrk(.S2)
        FnStrBrk = O
        Exit Function
    End With
End With
FnStrBrk = O
End Function

Sub FnStrBrk__Tst()
Dim Act$(): Act = FnStrBrk(" skldf dfk   kdf [df d] kdf df [  a ]  ")
Dim Exp$(): Exp = ApSy("skldf", "dfk", "kdf", "df d", "kdf", "df", "a")
AyChkEq Act, Exp
End Sub

Function FnStrIdxDic(FnStr) As Dictionary
Set FnStrIdxDic = AyIdxDic(FnStrBrk(FnStr))
End Function

Function FnStrLvc$(FnStr$)
Dim F$(): F = FnStrBrk(FnStr)
Dim J%
For J = 0 To UB(F)
    If StrIsNm(F(J)) Then F(J) = Quote(F(J), "[]")
Next
FnStrLvc = Join(F, ",")
End Function
