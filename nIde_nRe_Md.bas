Attribute VB_Name = "nIde_nRe_Md"
Option Compare Database
Option Explicit
Private X_MdLoc_S$
Private X_MdLoc_Lno&
Private X_MdLoc_Md As CodeModule

Sub MdLoc(S$, Optional A As CodeModule)


End Sub

Sub MdLocNxt()
Dim OLno&, OMd As CodeModule
    Set OMd = X_MdLoc_Md
    If IsNothing(OMd) Then Er "No last time X_MdLoc_Md"
    Dim A$(): A = MdLy(OMd, OLno)
    Dim I&: I = AyIdx_Contain(A, X_MdLoc_S)
    OLno = I
If OLno > 0 Then
    MdShwLno OLno, OMd
    X_MdLoc_Lno = OLno
End If
End Sub

Function MdLocRe(Patter$, Optional A As CodeModule) As Variant()


End Function

Sub MdLocReNxt()

End Sub
