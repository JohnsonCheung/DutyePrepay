Attribute VB_Name = "nOfc_App"
Option Compare Database
Option Explicit
Public Const AppxExt$ = ".xlam"
Public Const AppaExt$ = ".mda"
Public Const AppwExt$ = ".doca"
Public Const AppoExt$ = ".xlam"
Public Const ApppExt$ = ".ppta"
Public Const AppxExtNrm$ = ".xlsm"
Public Const AppaExtNrm$ = ".accdb"
Public Const AppwExtNrm$ = ".docx"
Public Const AppoExtNrm$ = ".pst"
Public Const ApppExtNrm$ = ".pptx"
Enum eOfcTy
    eXls
    eAcs
    eOlk
    eWrd
    ePpt
End Enum

Function AppExt$()
Dim O$
Select Case OfcTy
Case eXls: O = AppxExt
Case eAcs: O = AppaExt
Case eWrd: O = AppwExt
Case eOlk: O = AppoExt
Case ePpt: O = ApppExt
Case Else: Er "Unexpected {OfcTy}", OfcTy
End Select
AppExt = O
End Function

Function AppExtNrm$()
Dim O$
Select Case OfcTy
Case eXls: O = AppxExtNrm
Case eAcs: O = AppaExtNrm
Case eWrd: O = AppwExtNrm
Case eOlk: O = AppoExtNrm
Case ePpt: O = ApppExtNrm
Case Else: Er "Unexpected {OfcTy}", OfcTy
End Select
AppExtNrm = O
End Function

Function IsAppa() As Boolean
IsAppa = OfcTy = eAcs
End Function

Function IsAppo() As Boolean
IsAppo = OfcTy = eOlk
End Function

Function IsAppp() As Boolean
IsAppp = OfcTy = ePpt
End Function

Function IsAppw() As Boolean
IsAppw = OfcTy = eWrd
End Function

Function IsAppx() As Boolean
IsAppx = OfcTy = eXls
End Function

Function OfcTy() As eOfcTy
Dim O As eOfcTy
Select Case Application.Name
Case "Microsoft Excel":       O = eXls
Case "Microsoft Access":      O = eAcs
Case "Microsoft Word":        O = eWrd
Case "Microsoft Outlook":     O = eOlk
Case "Microsoft Power Point": O = ePpt
Case Else: Er "Unexpected {Application.Name}", Application.Name
End Select
OfcTy = O
End Function
