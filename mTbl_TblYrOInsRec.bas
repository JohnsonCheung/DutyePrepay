Attribute VB_Name = "mTbl_TblYrOInsRec"
Option Compare Database
Option Base 0
Option Explicit

Sub TblYrOInsRec()
'Aim: There is no current Yr record in table YrO, create one record in YrO
With CurrentDb.OpenRecordset("Select Yr from YrO where Yr=" & VBA.Year(Date) - 2000)
    If Not .EOF Then .Close: Exit Sub
    .Close
End With
SqlRun "Insert Into YrO (Yr) values (Year(Date())-2000)"
End Sub
