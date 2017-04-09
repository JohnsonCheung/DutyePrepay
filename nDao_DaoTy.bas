Attribute VB_Name = "nDao_DaoTy"
Option Compare Database
Option Explicit

Function DaoTyNew(DaoStr$) As DAO.DataTypeEnum
Dim O$
Select Case DaoStr
Case "DTE": O = dbDate
Case "INT": O = dbInteger
Case "LNG": O = dbLong
Case "DBL": O = dbDouble
Case "TXT": O = dbText
Case "SNG": O = dbSingle
Case "YES": O = dbBoolean
Case Else
    Er "Given {DaoStr} is not one of [DTE INT LNG DBL TXT SNG YES]", DaoStr
End Select
DaoTyNew = O
End Function

Function DaoTyNewByVbTy(VbTy As VbVarType) As DAO.DataTypeEnum
Dim O As DataTypeEnum
Select Case VbTy
Case VbVarType.vbArray:           GoTo X
Case VbVarType.vbBoolean:         O = dbBoolean
Case VbVarType.vbByte:            O = dbByte
Case VbVarType.vbCurrency:        O = dbCurrency
Case VbVarType.vbDataObject:      GoTo X
Case VbVarType.vbDate:            O = dbDate
Case VbVarType.vbDecimal:         O = dbDecimal
Case VbVarType.vbDouble:          O = dbDouble
Case VbVarType.vbEmpty:           GoTo X
Case VbVarType.vbError:           GoTo X
Case VbVarType.vbInteger:         O = dbInteger
Case VbVarType.vbLong:            O = dbLong
Case VbVarType.vbLongLong:        O = dbDecimal
Case VbVarType.vbNull:            GoTo X
Case VbVarType.vbObject:          GoTo X
Case VbVarType.vbSingle:          O = dbSingle
Case VbVarType.vbString:          O = dbText
Case VbVarType.vbUserDefinedType: GoTo X
Case VbVarType.vbVariant:         O = dbText
Case Else: Er "Given {VbTy} is invalid", VbTy
End Select
DaoTyNewByVbTy = O
Exit Function
X: Er "Given {VbTy} cannot convert to DataTypeEnum", VbTy
End Function

Function DaoTySqlStr$(A As DAO.DataTypeEnum)
Dim O$
Select Case A
Case dbDate: O = "Date"
Case dbInteger: O = "Int"
Case dbBoolean: O = "YesNo"
Case Else
    Er "DaoTySqlStr: Given {DataType} is not [dbDate dbInteger dbBoolean", A
End Select
DaoTySqlStr = O
End Function

Function DaoTyToStr$(A As DAO.DataTypeEnum)
Dim O$
Select Case A
Case dbDate:    O = "DTE"
Case dbInteger: O = "INT"
Case dbLong:    O = "LNG":
Case dbDouble:  O = "DBL":
Case dbText:    O = "TXT":
Case dbSingle:  O = "SNG"
Case dbBoolean: O = "YES"
Case Else
    Er "DaoTyToStr: Given {DataType} is not [dbDate dbInteger dbBoolean dbLong dbDouble dbText dbSingle]", A
End Select
DaoTyToStr = O
End Function
