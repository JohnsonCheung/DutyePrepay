Attribute VB_Name = "nIde_nPos_LCLC"
Option Compare Database
Option Explicit
Type LCLC
    L1 As Long
    C1 As Long
    L2 As Long
    C2 As Long
End Type

Function LCLCNew(L1, C1, L2, C2) As LCLC
Dim O As LCLC
With O
    .L1 = L1
    .L2 = L2
    .C1 = C1
    .C2 = C2
End With
LCLCNew = O
End Function
