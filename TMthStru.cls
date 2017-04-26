VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TMthStru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Private X_BEIdx() As Long
Public Brk As TMthBrk

Property Get BEIdx() As Long()
BEIdx = X_BEIdx
End Property

Property Set BEIdx(V)
X_BEIdx = V
End Property
