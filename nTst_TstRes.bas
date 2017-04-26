Attribute VB_Name = "nTst_TstRes"
Option Compare Database
Option Explicit

Function TstResFcsv$(MdNm$, Optional No%)
FbCurPth
End Function

Function TstResMdFcsv$(MdNm$, No%)
Dim N$
If No > 0 Then N = No
TstResMdFcsv = TstResMdPth(MdNm) & "F" & N & ".csv"
End Function

Function TstResMdPth$(MdNm$)
Dim O$
O = TstResPth & MdNm & "\"
PthEns O
TstResMdPth = O
End Function

Function TstResPth$()
Dim O$
O = FbCurPth & "TstRes\"
PthEns O
End Function
