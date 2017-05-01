Attribute VB_Name = "nIde_nVbe_Vbe"
Option Compare Database
Option Explicit

Function Pj(PjNm$, Optional A As Vbe) As vbproject
If PjNm = "" Then
    Set Pj = VbeNz(A).ActiveVBProject
    Exit Function
End If
Set Pj = VbeNz(A).VBProjects(PjNm)
End Function

Function VbeHasPj(Pj As vbproject, Optional A As Vbe) As Boolean
Dim P As vbproject
Dim O As LongPtr: O = ObjPtr(Pj)
For Each P In VbeNz(A).VBProjects
    If ObjPtr(P) = O Then VbeHasPj = True: Exit Function
Next
End Function

Function VbeLasPj(Optional A As VBIDE.Vbe) As vbproject
Dim V As Vbe: Set V = VbeNz(A)
Dim N%: N = V.VBProjects.Count
Set VbeLasPj = V.VBProjects(N)
End Function

Function VbeNz(Optional A As Vbe) As Vbe
If IsNothing(A) Then
    Set VbeNz = Application.Vbe
Else
    Set VbeNz = Vbe
End If
End Function

Function VbePjAy(Optional A As VBIDE.Vbe) As vbproject()
Dim O() As vbproject, I As vbproject
For Each I In VbeNz(A).VBProjects
    PushObj O, I
Next
VbePjAy = O
End Function

Function VbePjNy(Optional A As VBIDE.Vbe) As String()
VbePjNy = OyPrp_Nm(VbePjAy(A))
End Function
