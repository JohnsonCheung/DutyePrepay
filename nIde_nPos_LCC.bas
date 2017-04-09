Attribute VB_Name = "nIde_nPos_LCC"
Option Compare Database
Option Explicit
Type LCC
    L As Long
    C1 As Long
    C2 As Long
End Type

Sub LCCDmp(A As LCC)
Debug.Print LCCToStr(A)
End Sub

Function LCCNew(L, C1, C2) As LCC
Dim O As LCC
With O
    .L = L
    .C1 = C1
    .C2 = C2
End With
LCCNew = O
End Function

Function LCCToStr$(A As LCC)
With A
    LCCToStr = FmtQQ("L? C(? ?)", .L, .C1, .C2)
End With
End Function
