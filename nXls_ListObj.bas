Attribute VB_Name = "nXls_ListObj"
Option Compare Database
Option Explicit
Const C_Mod = "nXls_ListObj"

Function ListObjC(A As ListObject, C, Optional NoTot As Boolean) As Range
Dim O As Range
Set O = A.ListColumns(C).Range
If NoTot Then
    Set O = RgRR(O, 1, RgNRow(O) - 1)
End If
Set ListObjC = O
End Function

Sub ListObjC__Tst()
Dim Fx$
    Fx = PthTstRes(C_Mod) & "ListObj__Tst.xlsx"
Dim Wb As Workbook
    Set Wb = FxWb(Fx)
    'WbVis Wb
Dim ListObj As ListObject
    Set ListObj = WbFstWs(Wb).ListObjects(1)
Debug.Assert ListObjC(ListObj, 1).Address = "$B$2:$B$5"
Debug.Assert ListObjC(ListObj, "aa").Address = "$B$2:$B$5"
Debug.Assert ListObjC(ListObj, 1, NoTot:=True).Address = "$B$2:$B$4"
Debug.Assert ListObjC(ListObj, "aa", NoTot:=True).Address = "$B$2:$B$4"
WbCls Wb, NoSav:=True
End Sub

Function ListObjEntireC(A As ListObject, C) As Range
Set ListObjEntireC = ListObjC(A, C).EntireColumn
End Function

Function ListObjRC(A As ListObject, R, C) As Range
Set ListObjRC = RgRC(A.ListColumns(C).Range, 1, 1)
End Function
