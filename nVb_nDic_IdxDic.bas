Attribute VB_Name = "nVb_nDic_IdxDic"
Option Compare Database
Option Explicit

Sub AA2()
PkDicAsst__Tst
End Sub

Function IdxDicKy(IdxDic As Dictionary) As String()
Dim K$(): K = IdxDic.Keys
Dim J&, O$()
For J = 0 To UB(K)
    AySet O, CLng(K(J)), K
Next
End Function

Function PkDicAsst(IdxDic As Dictionary, Optional ErDrOpt)
ErAsst PkDicChk(IdxDic), ErDrOpt
End Function

Sub PkDicAsst__Tst()
PkDicAsst DicNew("A=0;B=1;C=2")
End Sub

Function PkDicChk(PkDic As Dictionary) As Dt
Dim OErDrAy()
    Dim K(): K = PkDic.Keys
    Dim U&: U = UB(K)
    Dim J&
    For J = 0 To U
        Stop
    Next
If AyIsEmpty(OErDrAy) Then Exit Function
Dim Dt As Dt: Dt = ErNewDrAy(OErDrAy)
PkDicChk = ErApd(Dt, "Given {PkDic} has error (Above)")
End Function
