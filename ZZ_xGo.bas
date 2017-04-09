Attribute VB_Name = "ZZ_xGo"
'Option Compare Text
'Option Explicit
'Option Base 0
'Const cMod$ = cLib & ".xGo"

Sub Go_Rec(Optional pWhere As AcRecord = acNext)
DoCmd.GoToRecord , , pWhere
End Sub

