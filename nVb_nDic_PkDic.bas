Attribute VB_Name = "nVb_nDic_PkDic"
Option Compare Database
Option Explicit

Sub AA2()
PkAsst__Tst
End Sub

Function PkAsst(Pk As Dictionary, Optional ErDrOpt)
ErAsst PkChk(Pk), ErDrOpt
End Function

Sub PkAsst__Tst()
PkAsst DicNew("A=0;B=1;C=2")
End Sub

Function PkChk(Pk As Dictionary) As Variant()
Dim Idx&()
Dim U&
    U = Pk.Count - 1
    ReDim Idx(U)
Dim OEr()
    Dim K(): K = Pk.Keys
    U = UB(K)
    Dim J&
    For J = 0 To U
        Stop
    Next
If AyIsEmpty(OEr) Then Exit Function
PkChk = ErApd(OEr, "Given {Pk} has error (Above)")
End Function

Function PkKey(Pk As Dictionary) As String()
Dim O$(), J&
Dim K(): K = Pk.Keys
For J = 0 To UB(O)
    O(Pk(K)) = K(J)
Next
PkKey = O
End Function

Function PkKy(Pk As Dictionary) As String()
Dim K$(): K = Pk.Keys
Dim J&, O$()
For J = 0 To UB(K)
    AySet O, CLng(K(J)), K
Next
End Function

Function PkNew(Ay$()) As Dictionary
Dim O As New Dictionary, J&
For J = 0 To UB(Ay)
    O.Add Ay(J), J
Next
Set PkNew = O
End Function
