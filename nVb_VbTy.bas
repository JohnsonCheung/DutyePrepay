Attribute VB_Name = "nVb_VbTy"
Option Compare Database
Option Explicit

Function VbTy(VbTyStr$) As VbVarType
Dim O$
Select Case VbTyStr
Case "AY":  O = VbVarType.vbArray
Case "YES": O = VbVarType.vbBoolean
Case "BYT": O = VbVarType.vbByte
Case "CUR": O = VbVarType.vbCurrency
Case "DTA": O = VbVarType.vbDataObject
Case "DTE": O = VbVarType.vbDate
Case "DEC": O = VbVarType.vbDecimal
Case "DBL": O = VbVarType.vbDouble
Case "EMP": O = VbVarType.vbEmpty
Case "ERR": O = VbVarType.vbError
Case "INT": O = VbVarType.vbInteger
Case "LNG": O = VbVarType.vbLong
Case "LLG": O = VbVarType.vbLongLong
Case "NUL": O = VbVarType.vbNull
Case "OBJ": O = VbVarType.vbObject
Case "SNG": O = VbVarType.vbSingle
Case "STR": O = VbVarType.vbString
Case "USR": O = VbVarType.vbUserDefinedType
Case "VAR": O = VbVarType.vbVariant
Case Else: Er "Given {VbTyStr} is in not {Vdt-VbTyStr}", VbTyStr, "AY YES BYT CUR DTA DTE DEC DBL EMP ERR INT LNG LLG NUL OBJ SNG STR USR VAR"
End Select
VbTy = O
End Function

Function VbTyByDaoTy(A As DAO.DataTypeEnum) As VbVarType
Dim O As VbVarType
Select Case A
Case DataTypeEnum.dbBigInt:         O = vbLongLong
Case DataTypeEnum.dbBoolean:        O = vbBoolean
Case DataTypeEnum.dbByte:           O = vbByte
Case DataTypeEnum.dbChar:           O = vbString
Case DataTypeEnum.dbComplexByte:    GoTo X
Case DataTypeEnum.dbComplexDecimal: GoTo X
Case DataTypeEnum.dbComplexDouble:  GoTo X
Case DataTypeEnum.dbComplexGUID:    GoTo X
Case DataTypeEnum.dbComplexInteger: GoTo X
Case DataTypeEnum.dbComplexLong:    GoTo X
Case DataTypeEnum.dbComplexSingle:  GoTo X
Case DataTypeEnum.dbComplexSingle:  GoTo X
Case DataTypeEnum.dbCurrency:       O = vbCurrency
Case DataTypeEnum.dbDate:           O = vbDate
Case DataTypeEnum.dbDecimal:        O = vbDecimal
Case DataTypeEnum.dbDouble:         O = vbDouble
Case DataTypeEnum.dbFloat:          O = vbSingle
Case DataTypeEnum.dbGUID:           GoTo X
Case DataTypeEnum.dbInteger:        O = vbInteger
Case DataTypeEnum.dbLong:           O = vbLong
Case DataTypeEnum.dbLongBinary:     GoTo X
Case DataTypeEnum.dbMemo:           O = vbString
Case DataTypeEnum.dbNumeric:        O = vbDouble
Case DataTypeEnum.dbSingle:         O = vbSingle
Case DataTypeEnum.dbText:           O = vbString
Case DataTypeEnum.dbTimeStamp:      O = vbDate
Case DataTypeEnum.dbVarBinary:      GoTo X
Case Else: Er "Given {Dao.DataTypeEnum} is invalid", A
End Select
VbTyByDaoTy = O
Exit Function
X: Er "Given {Dao.DataTypeEnum} cannot convert to VbVarType", A
End Function

Function VbTyStr$(T As VbVarType)
Dim O$
Select Case T
Case VbVarType.vbArray:           O = "AY"
Case VbVarType.vbBoolean:         O = "YES"
Case VbVarType.vbByte:            O = "BYT"
Case VbVarType.vbCurrency:        O = "CUR"
Case VbVarType.vbDataObject:      O = "DTA"
Case VbVarType.vbDate:            O = "DTA"
Case VbVarType.vbDecimal:         O = "DEC"
Case VbVarType.vbDouble:          O = "DBL"
Case VbVarType.vbEmpty:           O = "EMP"
Case VbVarType.vbError:           O = "ERR"
Case VbVarType.vbInteger:         O = "INT"
Case VbVarType.vbLong:            O = "LNG"
Case VbVarType.vbLongLong:        O = "LLG"
Case VbVarType.vbNull:            O = "NUL"
Case VbVarType.vbObject:          O = "OBJ"
Case VbVarType.vbSingle:          O = "SNG"
Case VbVarType.vbString:          O = "STR"
Case VbVarType.vbUserDefinedType: O = "USR"
Case VbVarType.vbVariant:         O = "VAR"
Case Else: Er "Given {VbTy} is as in VbVarType", T
End Select
VbTyStr = O
End Function
