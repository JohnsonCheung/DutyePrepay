Attribute VB_Name = "nAy_Oy"
Option Compare Database
Option Explicit

Function OyPrp(Oy, PrpNm$)
OyPrp = OyPrp_Into(Oy, PrpNm, EmptyVarAy)
End Function

Sub OyPrp__Tst()
Dim D As database
Set D = CurrentDb
Dim A() As TableDef:  A = TblAy(D)
Dim B$(): B = OyPrp_Nm(A)
AyBrw B
End Sub

Function OyPrp_Int(Oy, PrpNm$) As Integer()
OyPrp_Int = OyPrp_Into(Oy, PrpNm, EmptyIntAy)
End Function

Function OyPrp_Into(Oy, PrpNm$, OInto)
Erase OInto
Dim U&: U = UB(Oy)
If U >= 0 Then
    Dim J&
    ReDim OInto(U)
    For J = 0 To U
        OInto(J) = CallByName(Oy(J), PrpNm, VbGet)
    Next
End If
OyPrp_Into = OInto
End Function

Function OyPrp_Lng(Oy, PrpNm$) As Long()
OyPrp_Lng = OyPrp_Into(Oy, PrpNm, EmptyLngAy)
End Function

Function OyPrp_Nm(Oy) As String()
OyPrp_Nm = OyPrp_Str(Oy, "Name")
End Function

Function OyPrp_Str(Oy, PrpNm$) As String()
OyPrp_Str = OyPrp_Into(Oy, PrpNm, EmptySy)
End Function

Function OySelPrpEq(Oy, PrpNm$, EqVal)
Dim O: O = Oy: Erase O
Dim I
For Each I In Oy
    If CallByName(I, PrpNm, VbGet) = EqVal Then PushObj O, I
Next
OySelPrpEq = O
End Function

Function OySelPrpNe(Oy, PrpNm$, NeVal)
Dim O: O = Oy: Erase O
Dim I
For Each I In Oy
    If CallByName(I, PrpNm, VbGet) <> NeVal Then PushObj O, I
Next
OySelPrpNe = O
End Function

Sub OySetPrp(Oy, PrpNm$, V)
If AyIsEmpty(Oy) Then Exit Sub
Dim O
Dim Ty As VbCallType
For Each O In Oy
    Ty = IIf(IsObject(V), VbSet, VbLet)
    CallByName O, PrpNm, Ty, V  '<====
Next
End Sub

Function OySyPrp(Oy, PrpNm$) As String()
Dim O$()
Dim U&: U = UB(Oy)
If U >= 0 Then
    Dim J&
    ReDim O(U)
    For J = 0 To U
        O(J) = CallByName(Oy(J), PrpNm, VbGet)
    Next
End If
OySyPrp = O
End Function
