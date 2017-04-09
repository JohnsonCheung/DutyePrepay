Attribute VB_Name = "nTst_Pth"
Option Compare Database
Option Explicit

Function PthTstRes$(Optional MdNm$)
Dim O$
O = FbCurPth & "Res\" & MdNmNz(MdNm) & "\"
PthEnsAllSeg O
PthTstRes = O
End Function

Sub PthTstResBrw(Optional MdNm$)
PthBrw PthTstRes(MdNm)
End Sub
