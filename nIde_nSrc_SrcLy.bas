Attribute VB_Name = "nIde_nSrc_SrcLy"
Option Compare Database
Option Explicit

Function SrcLyDimLy(SrcLy$()) As String()
Dim A$(): A = AySel(SrcLy, "SrcLinIsDim")
Dim B$(): B = AyMapIntoSy(A, "DimLinTrim")
SrcLyDimLy = B
End Function

Function SrcLyDimNy(SrcLy$()) As String()
Dim A$(): A = SrcLyDimLy(SrcLy)
If AyIsEmpty(A) Then Exit Function
Dim O$(), I
For Each I In A
    PushAy O, DimLinNy(I)
Next
End Function

Function SrcLyOneContinueLin$(SrcLy$(), Idx&)
Dim O$()
Dim J&, L$
For J = Idx To Idx + 100
    L = SrcLy(J)
    If LasChr(L) <> "_" Then
        Push O, L
        SrcLyOneContinueLin = Join(O)
        Exit Function
    End If
    Push O, Trim(RmvLasChr(L))
Next
Er "SrcLyOneContinueLin: Impossible"
End Function
