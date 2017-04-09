Attribute VB_Name = "nDao_nFld_FldDclStr"
Option Compare Database
Option Explicit
Type FldDcl
    Nm As String
    Ty As DAO.DataTypeEnum
    L As Byte
End Type

Function FldDclBrk(FldDclStr) As FldDcl
Dim O As FldDcl
If Left(FldDclStr, 5) = "TEXT " Then
    O.Ty = dbText
    O.L = CByte(Mid(FldDclStr, 6))
    FldDclBrk = O
    Exit Function
End If

Select Case FldDclStr
Case "CURRENCY":     O.Ty = dbCurrency
Case "LONG", "AUTO": O.Ty = dbLong
Case "INT":          O.Ty = dbInteger
Case "BYTE":         O.Ty = dbByte
Case "DATE":         O.Ty = dbDate
Case "SINGLE":       O.Ty = dbSingle
Case "DOUBLE":       O.Ty = dbDouble
Case "MEMO":         O.Ty = dbMemo
Case "YESNO":        O.Ty = dbBoolean
Case Else: Er "FldDclStr} invalid", FldDclStr
End Select
FldDclBrk = O
End Function

Function FldDclFld(A As FldDcl) As Field
Dim O As Field
Set O = New Field
O.Name = A.Nm
O.Type = A.Ty
If A.Ty = dbText Then O.Size = A.L
Set FldDclFld = O
End Function

Function FldDclStrFld(FldDclStr$) As Field
Dim Dcl As FldDcl: Dcl = FldDclBrk(FldDclStr)
Set FldDclStrFld = FldDclFld(Dcl)
End Function
