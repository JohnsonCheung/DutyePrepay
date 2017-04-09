Attribute VB_Name = "nWrd_Doc"
Option Compare Database
Option Explicit

Sub DocCls(A As Word.Document, Optional pSav As Boolean = False, Optional pSilent As Boolean = False)
Dim W As Word.Application: Set W = Word.Application
W.DisplayAlerts = False
A.Save
End Sub

