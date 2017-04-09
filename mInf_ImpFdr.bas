Attribute VB_Name = "mInf_ImpFdr"
Option Compare Database
Option Explicit

Function ImpFdr$()
Dim O$: O = FbCurPth & "SAPDownloadExcel\"
PthEns O
ImpFdr = O
End Function

Sub ImpFdr__Tst()
PthBrw ImpFdr
End Sub
