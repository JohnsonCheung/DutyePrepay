Attribute VB_Name = "nDao_DaoTyScl"
Option Compare Database
Option Explicit

Function DaoTySclTyAy(DaoTySCL$) As DataTypeEnum()
Dim A$(): A = SclSy(DaoTySCL)
Dim U&: U = UB(A)
Dim O() As DataTypeEnum: ReSz O, U
Dim J%
For J = 0 To U
    O(J) = DaoTyNew(A(J))
Next
DaoTySclTyAy = O
End Function

Sub DaoTySclTyAy__Tst()
Const A = "TXT;INT;LNG;YES;DTE"
Dim Act() As DataTypeEnum: Act = DaoTySclTyAy(A)
Debug.Assert Sz(Act) = 5
Debug.Assert Act(0) = dbText
Debug.Assert Act(1) = dbInteger
Debug.Assert Act(2) = dbLong
Debug.Assert Act(3) = dbBoolean
Debug.Assert Act(4) = dbDate
End Sub
