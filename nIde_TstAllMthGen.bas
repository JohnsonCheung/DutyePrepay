Attribute VB_Name = "nIde_TstAllMthGen"
Option Compare Database
Option Explicit

Sub TstAllMthGenMd(Optional A As CodeModule)
TstAllMthRmv A
If Not TthIsInAnyMd(A) Then Exit Sub
Dim Lines$: Lines = TstAllMthLines(A)
End Sub

Function TstAllMthLines$(Optional A As CodeModule)
Dim TthNy$(): TthNy = TthNy_Md(A)
If AyIsEmpty(TthNy) Then Exit Function
Dim O$()
Push O, "Sub TstAll()"
PushAy O, TthNy
Push O, "End Sub"
TstAllMthLines = LyJn(O)
End Function

Sub TstAllMthLines__Tst()
Dim O$(), Lines$
Dim I, M As CodeModule
For Each I In PjMdAy
    Set M = I
    Lines = TstAllMthLines(M)
    If Lines <> "" Then
        Push O, MdNm(M)
        Push O, Lines
        Push O, ""
    End If
Next
AyBrw O
End Sub

Sub TstAllMthRmv(Optional A As CodeModule)
MthRmv "TstAll", A
End Sub
