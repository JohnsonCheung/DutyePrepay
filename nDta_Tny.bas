Attribute VB_Name = "nDta_Tny"
Option Compare Database
Option Explicit

Function TnyOptDic(TnyOpt, TblUB&) As Dictionary
Dim J&
Dim O As New Dictionary
    If IsMissing(TnyOpt) Then
        For J = 0 To TblUB
            O.Add "Tbl" & J, J
        Next
    ElseIf TypeName(TnyOpt) = "Dictionary" Then
        Set O = TnyOpt
    ElseIf VarIsStr(TnyOpt) Then
        Set O = FnStrIdxDic(TnyOpt)
    ElseIf VarIsSy(TnyOpt) Then
        Dim Tn, I&
        For Each Tn In TnyOpt
            O.Add Tn, I
            I = I + 1
        Next
    Else
        Er "Given TnyOpt-{Ty} should be [Missing | Dic | String | String array]", TypeName(TnyOpt)
    End If
Set TnyOptDic = O
End Function
