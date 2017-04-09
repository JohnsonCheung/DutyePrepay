Attribute VB_Name = "nIde_nConstSrcLin_Md"
Option Compare Database
Option Explicit
Const C_Mod = "nIde_nConstSrcLin_Md"

Function MdIsUsingConstMod(Optional A As CodeModule) As Boolean
Dim L$: L = MdBdyLines(A)
MdIsUsingConstMod = StrHas(L, "C_Mod")
End Function

