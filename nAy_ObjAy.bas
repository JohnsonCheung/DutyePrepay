Attribute VB_Name = "nAy_ObjAy"
Option Compare Database
Option Explicit

Function ObjAyNy(ObjAy) As String()
ObjAyNy = ObjAySyPrp(ObjAy, "Name")
End Function

Function ObjAyPrp(ObjAy, PrpNm$, OAy)
Erase OAy
Dim U&: U = UB(ObjAy)
If U >= 0 Then
    Dim J&
    ReDim OAy(U)
    For J = 0 To U
        OAy(J) = CallByName(ObjAy(J), PrpNm, VbGet)
    Next
End If
ObjAyPrp = OAy
End Function

Sub ObjAyPrp__Tst()
Dim D As database
Set D = CurrentDb
Dim A() As TableDef:  A = DbTblDefAy(D)
Dim B$(): B = ObjAyPrp(A, "Name", B)
AyBrw B
End Sub

Sub ObjAySetPrp(ObjAy, PrpNm$, V)
If AyIsEmpty(ObjAy) Then Exit Sub
Dim O
Dim Ty As VbCallType
For Each O In ObjAy
    Ty = IIf(IsObject(V), VbSet, VbLet)
    CallByName O, PrpNm, Ty, V  '<====
Next
End Sub

Function ObjAyStrPrp(ObjAy, PrpNm$)
Dim O$()
Dim U&: U = UB(ObjAy)
If U >= 0 Then
    Dim J&
    ReDim O(U)
    For J = 0 To U
        O(J) = CallByName(ObjAy(J), PrpNm, VbGet)
    Next
End If
ObjAyStrPrp = O
End Function

Function ObjAySyPrp(ObjAy, PrpNm$) As String()
Dim O$()
Dim U&: U = UB(ObjAy)
If U >= 0 Then
    Dim J&
    ReDim O(U)
    For J = 0 To U
        O(J) = CallByName(ObjAy(J), PrpNm, VbGet)
    Next
End If
ObjAySyPrp = O
End Function
