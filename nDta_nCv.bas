Attribute VB_Name = "nDta_nCv"
Option Compare Text
Option Explicit
Option Base 0
Enum eSimTy
    eSimNum = 1
    eSimStr = 2
    eSimBool = 3
    eSimDte = 4
    eSimOth = 5
End Enum

Function DaoTySzFmSql(Sql$) As Variant()
Dim OTy As DAO.DataTypeEnum
Dim OSz%
If Left(Sql, 5) = "TEXT " Then
    OTy = dbText
    OSz = CByte(Mid(Sql, 6))
    Exit Function
End If
Select Case Sql
Case "CURRENCY":     OTy = dbCurrency
Case "LONG", "AUTO": OTy = dbLong
Case "INT":          OTy = dbInteger
Case "BYTE":         OTy = dbByte
Case "DATE":         OTy = dbDate
Case "SINGLE":       OTy = dbSingle
Case "DOUBLE":       OTy = dbDouble
Case "MEMO":         OTy = dbMemo
Case "YESNO":        OTy = dbBoolean
Case Else: Er "Invalid Fld-Dcl-Sql-Phrase"
End Select
DaoTySzFmSql = Array(OTy, OSz)
End Function

Function DaoTySql$(A As DAO.DataTypeEnum, Optional Sz%)
Dim O$
Select Case A
Case DAO.DataTypeEnum.dbBigInt _
   , DAO.DataTypeEnum.dbLong
                                    O = "Long"
Case DAO.DataTypeEnum.dbByte
                                    O = "Byte"
Case DAO.DataTypeEnum.dbCurrency _
   , DAO.DataTypeEnum.dbDecimal
                                    O = "Currency"

Case DAO.DataTypeEnum.dbDouble _
   , DAO.DataTypeEnum.dbFloat _
   , DAO.DataTypeEnum.dbNumeric
                                    O = "Double"
Case DAO.DataTypeEnum.dbInteger
                                    O = "Integer"
Case DAO.DataTypeEnum.dbSingle
                                    O = "Single"
Case DAO.DataTypeEnum.dbMemo
                                    O = "Memo"
Case DAO.DataTypeEnum.dbChar _
    , DAO.DataTypeEnum.dbText
                                    O = "Text " & Sz
Case DAO.DataTypeEnum.dbBoolean
                                    O = "YesNo"
Case DAO.DataTypeEnum.dbDate _
    , DAO.DataTypeEnum.dbTime _
    , DAO.DataTypeEnum.dbTimeStamp
                                    O = "Date"
Case Else
                                    O = "Unexpect DaoTy (" & A & ")"
End Select
DaoTySql = O
End Function

Function VarByStr(Str$, Ty As eSimTy)
Dim O
Select Case Ty
Case eSimBool: O = CBool(Str)
Case eSimDte: O = CDate(Str): If O < #1/1/1990# Then ss.A 1, "Date is less <1990/1/1": GoTo E
Case eSimNum: O = CDbl(Str)
Case eSimOth: ss.A 2, "SimTy(Other) is not handled", "Vraw", Str: GoTo E
Case eSimStr: O = Str
Case Else: ss.A 3, "Unexpected SimTy", "Vraw,SimTy", Str, Ty: GoTo E
End Select
VarByStr = O
Exit Function
R: ss.R
E:
End Function

Function DaoTy2QChr$(A As DatabaseTypeEnum)
DaoTy2QChr = SimTyQChr(DaoTyToSim(A))
End Function

Function DaoTyToSim(pDaoTy As DAO.DataTypeEnum) As eSimTy
Select Case pDaoTy
Case DAO.DataTypeEnum.dbBigInt _
    , DAO.DataTypeEnum.dbByte _
    , DAO.DataTypeEnum.dbCurrency _
    , DAO.DataTypeEnum.dbDecimal _
    , DAO.DataTypeEnum.dbDouble _
    , DAO.DataTypeEnum.dbFloat _
    , DAO.DataTypeEnum.dbInteger _
    , DAO.DataTypeEnum.dbLong _
    , DAO.DataTypeEnum.dbNumeric _
    , DAO.DataTypeEnum.dbSingle
                                    DaoTyToSim = eSimNum
Case DAO.DataTypeEnum.dbChar _
    , DAO.DataTypeEnum.dbMemo _
    , DAO.DataTypeEnum.dbText
                                    DaoTyToSim = eSimStr
Case DAO.DataTypeEnum.dbBoolean
                                    DaoTyToSim = eSimBool
Case DAO.DataTypeEnum.dbDate _
    , DAO.DataTypeEnum.dbTime _
    , DAO.DataTypeEnum.dbTimeStamp
                                    DaoTyToSim = eSimDte
Case Else
                                    DaoTyToSim = eSimOth
End Select
End Function


Function SimTyToStr$(A As eSimTy)
Dim O$
Select Case A
Case eSimNum: O = "Num"
Case eSimStr: O = "Chr"
Case eSimDte: O = "Dte"
Case eSimBool: O = "Bool"
Case eSimOth: O = "Oth"
End Select
SimTyToStr = O
End Function

Function VarSimTy(V) As eSimTy
VarSimTy = VbTySim(VarType(V))
End Function
Function SimTyQChr$(A As eSimTy)
Select Case A
Case eSimStr: SimTyQChr = CtSngQ
Case eSimDte: SimTyQChr = "#"
Case Else: SimTyQChr = ""
End Select
End Function

Function SimTyChr$(A As eSimTy)
Select Case A
Case eSimBool: SimTyChr = "B"
Case eSimDte: SimTyChr = "D"
Case eSimNum: SimTyChr = "N"
Case eSimOth: SimTyChr = "O"
Case eSimStr: SimTyChr = "S"
Case Else: SimTyChr = "?"
End Select
End Function


Function VbTySim(pVbVarTyp As VbVarType) As eSimTy
Select Case pVbVarTyp
Case VBA.VbVarType.vbByte _
    , VBA.VbVarType.vbCurrency _
    , VBA.VbVarType.vbDecimal _
    , VBA.VbVarType.vbDouble _
    , VBA.VbVarType.vbInteger _
    , VBA.VbVarType.vbLong _
    , VBA.VbVarType.vbSingle
                                    VbTySim = eSimNum
Case VBA.VbVarType.vbString
                                    VbTySim = eSimStr
Case VBA.VbVarType.vbBoolean
                                    VbTySim = eSimBool
Case VBA.VbVarType.vbDate
                                    VbTySim = eSimDte
Case Else
                                    VbTySim = eSimOth
End Select
End Function

