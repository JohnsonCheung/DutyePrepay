VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MthPrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Public Mth As Mth
Private X_PrmAy() As Prm

Property Get PrmAy() As Prm()
PrmAy = X_PrmAy
End Property

Sub SetPrmAy(V() As Prm)
X_PrmAy = V
End Sub
