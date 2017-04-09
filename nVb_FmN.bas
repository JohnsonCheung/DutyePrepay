Attribute VB_Name = "nVb_FmN"
Option Compare Database
Option Explicit

Function FmNNew(Fm&, N&) As Long()
If Fm < 0 Then Er "FmNNew: {Fm} < 0", Fm
If N < 0 Then Er "FmNNew: {N} < 0", N
FmNNew = ApLngAy(Fm, N)
End Function
