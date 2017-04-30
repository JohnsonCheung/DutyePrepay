Attribute VB_Name = "nIde_FctTy"
Option Compare Database
Option Explicit
Public Enum eFctTy
    eSub = 1
    eFct = 2
    eGet = 3
    eLet = 4
    eSet = 5
End Enum

Function FctTyToStr$(A As eFctTy)
Dim O$
Select Case A
Case eFctTy.eFct: O = "Function "
Case eFctTy.eSub: O = "Sub "
Case eFctTy.eGet: O = "Property Get "
Case eFctTy.eLet: O = "o Let "
Case eFctTy.eSet: FctTyToStr = "Property Set "
Case Else: O = "Unknown TypFct(" & pTypFct & ")"
End Select
FctTyToStr = O
End Function

