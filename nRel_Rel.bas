Attribute VB_Name = "nRel_Rel"
Option Compare Database
Option Explicit

Sub RelAsstVdt(A As Dictionary)
ErAsst RelChkVdt(A), "Given Relation is invalid"
End Sub

Sub RelBrw(A As Dictionary)
AyBrw RelLy(A), WithIdx:=True
End Sub

Sub RelBrw__Tst()
RelBrw RelSample1
End Sub

Function RelChd(A As Dictionary) As String()
If RelIsEmpty(A) Then Exit Function
Dim K, O$()
For Each K In A.Keys
    PushAyNoDup O, A(K)
Next
RelChd = O
End Function

Function RelChkVdt(A As Dictionary) As Variant()

End Function

Function RelIsCyc(A As Dictionary) As Boolean
Dim ParAy$(): ParAy = RelParAy(A)
Dim ChdAy, I
For Each ChdAy In A.Items
    For Each I In ChdAy
        If AyHas(ParAy, I) Then RelIsCyc = True: Exit Function
    Next
Next
End Function

Function RelIsEmpty(A As Dictionary) As Boolean
RelIsEmpty = A.Count = 0
End Function

Function RelIsSngRoot(A As Dictionary) As Boolean
RelIsSngRoot = Sz(RelRootAy(A)) = 1
End Function

Function RelItmAy(A As Dictionary) As String()
Dim O$(), Chd
O = RelParAy(A)
For Each Chd In A.Values
    PushAyNoDup O, Chd
Next
RelItmAy = O
End Function

Function RelItmAy_Cyc(A As Dictionary) As String()
Dim O$()
Dim ParAy$(): ParAy = RelParAy(A)
Dim ChdAy, I
For Each ChdAy In A.Items
    For Each I In ChdAy
        If AyHas(ParAy, I) Then Push O, I
    Next
Next
RelItmAy_Cyc = O
End Function

Function RelLeafAy(A As Dictionary) As String()

End Function

Function RelLy(A As Dictionary) As String()
RelLy = AyMapInto(RelNodAy(A), ApSy, "NodToStr")
End Function

Function RelNew(RelLy$()) As Dictionary
Dim J&, O As New Dictionary, Nod()
For J = 0 To UB(RelLy)
    Nod = NodNew(RelLy(J))
    O.Add Nod(0), Nod(1)
Next
Set RelNew = O
End Function

Sub RelNew__Tst()
Dim A As Dictionary: Set A = RelSample1
Debug.Assert A.Count = 3
Dim K: K = A.Keys
Debug.Assert Sz(K) = 3
Debug.Assert K(0) = "A"
Debug.Assert K(1) = "B"
Debug.Assert K(2) = "H"
Debug.Assert AyIsEq(A("A"), ApSy("B", "C", "D"))
Debug.Assert AyIsEq(A("B"), ApSy("E", "F", "G"))
Debug.Assert AyIsEq(A("H"), ApSy("I", "J", "K", "L"))
End Sub

Function RelNodAy(A As Dictionary) As Variant()
If RelIsEmpty(A) Then Exit Function
Dim O(), K
For Each K In A.Keys
    Push O, Array(K, A(K))
Next
RelNodAy = O
End Function

Function RelParAy(A As Dictionary) As String()
If DicIsEmpty(A) Then Exit Function
RelParAy = A.Keys
End Function

Function RelRootAy(A As Dictionary) As String()

End Function

Function RelSample1() As Dictionary
Set RelSample1 = RelNew(SplitVBar("A:B C D|B:E F G|H:I J K L"))
End Function
