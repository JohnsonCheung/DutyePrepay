Attribute VB_Name = "nIde_nConstSrcLin_ConstModLin"
Option Compare Database
Option Explicit
Const C_Mod$ = "nIde_nConstSrcLin_ConstModLin"

Sub ConstModEns(Optional A As CodeModule)
Dim Md As CodeModule: Set Md = MdNz(A)
Dim ExpLin$: ExpLin = ConstModExpLin(Md)
Dim ActLin$: ActLin = ConstModLin(Md)
If MdIsUsingConstMod(Md) Then
    If ExpLin = ActLin Then
        Debug.Print "[Const C_Mod] Line no change " & MdNm(Md)
        Exit Sub
    End If
    ConstModRmv Md
    Md.InsertLines MdLnoAftOptLin(Md), ExpLin
    Debug.Print "[Const C_Mod] Line replaced  " & MdNm(Md)
    
Else
    If ActLin = "" Then
        Debug.Print "[Const C_Mod] Line not using " & MdNm(Md)
        Exit Sub
    End If
    ConstModRmv Md
    Debug.Print "[Const C_Mod] Line removed   " & MdNm(Md)
End If
End Sub

Sub ConstModEnsPj(Optional A As vbproject)
AyEachEle PjMdAy(A), "ConstModEns"
End Sub

Function ConstModExpLin$(Optional A As CodeModule)
ConstModExpLin = FmtQQ("Const C_Mod$ = ""?""", MdNm(A))
End Function

Function ConstModLin$(Optional A As CodeModule)
Dim L&: L = ConstModLno(A)
ConstModLin = MdOneLin(L, A)
End Function

Function ConstModLno&(A As CodeModule)
Dim Ly$(): Ly = MdDclLy(A)
Dim J&
For J = 0 To UB(Ly)
    If IsPfx(Ly(J), "Const C_Mod") Then ConstModLno = 1 + J: Exit Function
Next
End Function

Sub ConstModRmv(A As CodeModule)
Dim Md As CodeModule: Set Md = MdNz(A)
Dim L&: L = ConstModLno(Md)
If L = 0 Then Exit Sub
MdRmvOneLin L, Md
End Sub
