Attribute VB_Name = "nIde_nVbe_Vbe"
Option Compare Database
Option Explicit

Function NzVbe(Optional A As Vbe) As Vbe
If IsNothing(A) Then
    Set NzVbe = Application.Vbe
Else
    Set NzVbe = Vbe
End If
End Function

Function Pj(PjNm$, Optional A As Vbe) As vbproject
If PjNm = "" Then
    Set Pj = NzVbe(A).ActiveVBProject
    Exit Function
End If
Set Pj = NzVbe(A).VBProjects(PjNm)
End Function

Function VbeHasPj(Pj As vbproject, Optional A As Vbe) As Boolean
Dim P As vbproject
Dim O As LongPtr: O = ObjPtr(Pj)
For Each P In NzVbe(A).VBProjects
    If ObjPtr(P) = O Then VbeHasPj = True: Exit Function
Next
End Function

Function VbeLasPj(Optional A As VBIDE.Vbe) As vbproject
Dim V As Vbe: Set V = NzVbe(A)
Dim N%: N = V.VBProjects.Count
Set VbeLasPj = V.VBProjects(N)
End Function

Function VbePjAy(Optional A As VBIDE.Vbe) As vbproject()
Dim O() As vbproject, I As vbproject
For Each I In NzVbe(A).VBProjects
    PushObj O, I
Next
VbePjAy = O
End Function

Function VbePjNy(Optional A As VBIDE.Vbe) As String()
VbePjNy = ObjAyPrp(VbePjAy(A), "Name", ApSy)
End Function
