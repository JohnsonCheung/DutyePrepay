Attribute VB_Name = "nAcs_AcsPrp"
Option Compare Database
Option Explicit

Function AcsFrmOy(Optional NmStr$, Optional A As Access.Application) As AccessObject()
AcsFrmOy = CollOy(AppaNz(A).CurrentProject.AllForms, EmptyAcsOy)
End Function

Function AcsMdOy(Optional NmStr$, Optional A As Access.Application) As AccessObject()
AcsMdOy = CollOy(AppaNz(A).CurrentData.AllModules, EmptyAcsOy)
End Function

Function AcsOy(Optional NmStr$, Optional Ty As AcObjectType, Optional A As Access.Application) As Access.AccessObject()
Dim Acs As Access.Application: Set Acs = Appa(A)
Select Case Ty
Case 0:
Case AcObjectType.acForm: AcsOy = AcsFrmOy(NmStr, A)
Case AcObjectType.acQuery: AcsOy = AcsQryOy(NmStr, A)
Case Else: Er "AcsOy: Invalid {Ty}", Ty
End Select
End Function

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

Function AcsQryOy(Optional NmStr$, Optional A As Access.Application) As AccessObject()
AcsQryOy = CollOy(AppaNz(A).CurrentData.AllQueries, EmptyAcsOy)
End Function

Sub AcsQryOy__Tst()
Dim O() As AccessObject
O = AcsQryOy
AyBrw OyPrp_Nm(O)
End Sub
