Attribute VB_Name = "nAcs_AcsPrp"
Option Compare Database
Option Explicit

Sub AcsPrpSet(ObjNm, ObjTy As AcObjectType, PrpNm$, V)
'Dim mPrp As DAO.Property:
'Select Case ObjTy
'Case Access.AcObjectType.acQuery
'    On Error GoTo Er1
'    CurrentDb.QueryDefs(pNm).Properties(PrpNm).Value = V
'Case acForm, acModule, acReport, acTable
'    Dim mNmTypObj$:  mNmTypObj = ToStr_TypObj(ObjTy)
'    On Error GoTo Er2
'    CurrentDb.Containers(mNmTypObj).Documents(pNm).Properties(PrpNm).Value = V
'Case Else
'    ss.A 1, "Given TypObj is not supported", , "Supported Types", "acQuery, acForm, acModule, acReport, acTable": GoTo E
'End Select
End Sub

Function AcsPrpSet__Tst()
'Debug.Print Set_Prp("1Rec", acForm, "Description", "yy")
End Function
