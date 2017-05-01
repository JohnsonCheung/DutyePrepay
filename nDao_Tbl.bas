Attribute VB_Name = "nDao_Tbl"
Option Compare Database
Option Explicit

Function Tbl(T, Optional A As database) As TableDef
Set Tbl = DbNz(A).TableDefs(T)
End Function

Function TblAy(Optional A As database) As TableDef()
Dim O() As TableDef
Dim I As TableDef, J%
For Each I In DbNz(A).TableDefs
    PushObj O, I
Next
TblAy = O
End Function
