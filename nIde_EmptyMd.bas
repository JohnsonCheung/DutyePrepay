Attribute VB_Name = "nIde_EmptyMd"
Option Compare Database
Option Explicit

Function EmptyMdNy(Optional A As vbproject) As String()
Dim EmptyMd() As CodeModule: EmptyMd = AySel(PjMdAy(A), "MdIsEmpty")
EmptyMdNy = AyMapInto(EmptyMd, ApSy, "MdNm")
End Function

Sub EmptyMdNy__Tst()
AyBrw EmptyMdNy
End Sub

Sub EmptyMdRmv(Optional A As vbproject)

End Sub
