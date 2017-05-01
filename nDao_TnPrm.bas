Attribute VB_Name = "nDao_TnPrm"
Option Compare Database
Option Explicit

Function TnPrmToTny(TnPrm, A As database) As String()
Dim O$()
    If IsMissing(TnPrm) Then
        O = DbTny(A)
    ElseIf VarIsStr(TnPrm) Then
        O = NmBrk(TnPrm)
    ElseIf VarIsSy(TnPrm) Then
        O = TnPrm
    ElseIf VarIsDic(TnPrm) Then
        Dim D As Dictionary: Set D = TnPrm
        O = AySy(D.Keys)
    Else
        Er "Given TnPrm has unexpected {Ty}.  Exp-Ty is [Missing | Str | StrAy | Dic]", TypeName(TnPrm)
    End If
TnPrmToTny = O
End Function
