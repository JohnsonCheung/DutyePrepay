Attribute VB_Name = "nVb_Opt"
Option Compare Database
Option Explicit
Type OptDte
    Som As Boolean
    Dte As Date
End Type
Type OptStr
    Som As Boolean
    Str As String
End Type

Type OptCur
    Som As Boolean
    Cur As Currency
End Type

Type OptInt
    Som As Boolean
    Int As Integer
End Type

Type OptLng
    Som As Boolean
    Lng As Long
End Type

Type OptV
    Som As Boolean
    V As Variant
End Type

Function OptSy(SyOpt) As String()
If VarIsStr(SyOpt) Then
    OptSy = NmBrk(SyOpt)
ElseIf VarIsSy(SyOpt) Then
    OptSy = SyOpt
ElseIf IsMissing(SyOpt) Then
    Exit Function
ElseIf VarIsDic(SyOpt) Then
    OptSy = AySy(SyOpt.Keys)
Else
    Er "OptSy: Given SyOpt-{Ty} is not [Str StrAy Dic Missing]", TypeName(SyOpt)
End If
End Function

Function OptVNew(V) As OptV
OptVNew.V = V
OptVNew.Som = True
End Function
