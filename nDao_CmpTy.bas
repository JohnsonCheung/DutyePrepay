Attribute VB_Name = "nDao_CmpTy"
Option Compare Database
Option Explicit

Function CmpTyToStr$(CmpTy As VBIDE.vbext_ComponentType)
Select Case CmpTy
Case VBIDE.vbext_ComponentType.vbext_ct_ActiveXDesigner:    CmpTyToStr = "ActX"
Case VBIDE.vbext_ComponentType.vbext_ct_ClassModule:        CmpTyToStr = "Class"
Case VBIDE.vbext_ComponentType.vbext_ct_Document:           CmpTyToStr = "Doc"
Case VBIDE.vbext_ComponentType.vbext_ct_MSForm:             CmpTyToStr = "Frm"
Case VBIDE.vbext_ComponentType.vbext_ct_StdModule:          CmpTyToStr = "Mod"
Case Else: CmpTyToStr = "Unknow(" & CmpTy & ")"
End Select
End Function
