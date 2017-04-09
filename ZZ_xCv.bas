Attribute VB_Name = "ZZ_xCv"
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xCv"

Function Cv_RnoColRge2Adr$(pColRge$, pRow&, Optional pNRow& = 1)
Dim P As Byte: P = InStr(pColRge, ":")
If P = 0 Then Cv_RnoColRge2Adr = pColRge & pRow & ":" & pColRge & pRow + pNRow - 1: Exit Function
Cv_RnoColRge2Adr = Left(pColRge, P - 1) & pRow & ":" & Mid(pColRge, P + 1) & pRow + pNRow - 1
End Function

Function Cv_Tbl2Tbl(pNmtFm$, pNmtTo$, pLm$) As Boolean
'Aim: Transform {pNmtFm} into {pNmtTo} by using {pLm}
Const cSub$ = "Tbl2Tbl"
On Error GoTo R
If Not IsTbl(pNmtFm) Then ss.A 1, "pNmtFm not exist", , "pNmtFm,pNmtTo", pNmtFm, pNmtTo
If Dlt_Tbl(pNmtTo) Then ss.A 2: GoTo E
Dim mAm() As tMap: mAm = Get_Am_ByLm(pLm)
Dim mAyFm$(): If Cpy_AmF1_ToAy(mAyFm, mAm) Then ss.A 4: GoTo E
Dim mRs As DAO.Recordset: Set mRs = CurrentDb.TableDefs(pNmtFm).OpenRecordset
If Chk_Struct_Rs(mRs, Join(mAyFm, CtComma)) Then ss.A 5: GoTo E
Dim mSel$: mSel = ToStr_Am(mAm, " as ", "[]", "[]")
Dim mSql$: mSql = Fmt_Str("Select {0} into {1} from {2}", mSel, Rmv_SqBkt(pNmtTo), Rmv_SqBkt(pNmtFm))
If Run_Sql(mSql) Then ss.A 6: GoTo E
Exit Function
R: ss.R
E: Cv_Tbl2Tbl = True: ss.B cSub, cMod, "pNmtFm,pNmtTo,pLm", pNmtFm, pNmtTo, pLm
End Function

Function Cv_TypDAO_FmFldDcl(oTypDao As DAO.DataTypeEnum, oLen As Byte, pFldDcl$) As Boolean
Const cSub$ = "Cv_TypDAO_FmFldDcl"
On Error GoTo R
If Left(pFldDcl, 5) = "TEXT " Then
    oTypDao = dbText
    oLen = CByte(Mid(pFldDcl, 6))
    Exit Function
End If
Select Case pFldDcl
Case "CURRENCY":     oTypDao = dbCurrency
Case "LONG", "AUTO": oTypDao = dbLong
Case "INT":          oTypDao = dbInteger
Case "BYTE":         oTypDao = dbByte
Case "DATE":         oTypDao = dbDate
Case "SINGLE":       oTypDao = dbSingle
Case "DOUBLE":       oTypDao = dbDouble
Case "MEMO":         oTypDao = dbMemo
Case "YESNO":        oTypDao = dbBoolean
Case Else: ss.A 1, "Invalid pFldDcl": GoTo E
End Select
Exit Function
R: ss.R
E: Cv_TypDAO_FmFldDcl = True: ss.B cSub, cMod, "pFldDcl", pFldDcl
End Function

Function Cv_TypDAO_ToFldDcl$(pTypDAO As DAO.DataTypeEnum, Optional pLen As Byte)
Select Case pTypDAO
Case DAO.DataTypeEnum.dbBigInt _
   , DAO.DataTypeEnum.dbLong
                                    Cv_TypDAO_ToFldDcl = "Long"
Case DAO.DataTypeEnum.dbByte
                                    Cv_TypDAO_ToFldDcl = "Byte"
Case DAO.DataTypeEnum.dbCurrency _
   , DAO.DataTypeEnum.dbDecimal
                                    Cv_TypDAO_ToFldDcl = "Currency"

Case DAO.DataTypeEnum.dbDouble _
   , DAO.DataTypeEnum.dbFloat _
   , DAO.DataTypeEnum.dbNumeric
                                    Cv_TypDAO_ToFldDcl = "Double"
Case DAO.DataTypeEnum.dbInteger
                                    Cv_TypDAO_ToFldDcl = "Integer"
Case DAO.DataTypeEnum.dbSingle
                                    Cv_TypDAO_ToFldDcl = "Single"
Case DAO.DataTypeEnum.dbMemo
                                    Cv_TypDAO_ToFldDcl = "Memo"
Case DAO.DataTypeEnum.dbChar _
    , DAO.DataTypeEnum.dbText
                                    Cv_TypDAO_ToFldDcl = "Text " & pLen
Case DAO.DataTypeEnum.dbBoolean
                                    Cv_TypDAO_ToFldDcl = "YesNo"
Case DAO.DataTypeEnum.dbDate _
    , DAO.DataTypeEnum.dbTime _
    , DAO.DataTypeEnum.dbTimeStamp
                                    Cv_TypDAO_ToFldDcl = "Date"
Case Else
                                    Cv_TypDAO_ToFldDcl = "Unexpect TypDAO (" & pTypDAO & ")"
End Select
End Function

Function Cv_Vraw2Val(oVal, pVraw$, pTypSim As eTypSim) As Boolean
Const cSub$ = "Vraw2Val"
On Error GoTo R
Select Case pTypSim
Case eTypSim_Bool: oVal = CBool(pVraw)
Case eTypSim_Dte: oVal = CDate(pVraw): If oVal < #1/1/1990# Then ss.A 1, "Date is less <1990/1/1": GoTo E
Case eTypSim_Num: oVal = CDbl(pVraw)
Case eTypSim_Oth: ss.A 2, "TypSim(Other) is not handled", "pVraw", pVraw: GoTo E
Case eTypSim_Str: oVal = pVraw
Case Else: ss.A 3, "Unexpected TypSim", "pVraw,TypSim", pVraw, pTypSim: GoTo E
End Select
Exit Function
R: ss.R
E: Cv_Vraw2Val = True: ss.B cSub, cMod, "pVraw,pTypSim", pVraw, pTypSim
End Function

Function DaoTy2QChr$(pTypDAO As DatabaseTypeEnum)
DaoTy2QChr = TypSim2QChr(DaoTyToSim(pTypDAO))
End Function

Function DaoTyToSim(pTypDAO As DAO.DataTypeEnum) As eTypSim
Select Case pTypDAO
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
                                    DaoTyToSim = eTypSim_Num
Case DAO.DataTypeEnum.dbChar _
    , DAO.DataTypeEnum.dbMemo _
    , DAO.DataTypeEnum.dbText
                                    DaoTyToSim = eTypSim_Str
Case DAO.DataTypeEnum.dbBoolean
                                    DaoTyToSim = eTypSim_Bool
Case DAO.DataTypeEnum.dbDate _
    , DAO.DataTypeEnum.dbTime _
    , DAO.DataTypeEnum.dbTimeStamp
                                    DaoTyToSim = eTypSim_Dte
Case Else
                                    DaoTyToSim = eTypSim_Oth
End Select
End Function

Function TypSim2QChr$(pTypSim As eTypSim)
Select Case pTypSim
Case eTypSim_Str: TypSim2QChr = CtSngQ
Case eTypSim_Dte: TypSim2QChr = "#"
Case Else: TypSim2QChr = ""
End Select
End Function

Function TypSimToChr$(pTyp As eTypSim)
Select Case pTyp
Case eTypSim_Bool: TypSimToChr = "B"
Case eTypSim_Dte: TypSimToChr = "D"
Case eTypSim_Num: TypSimToChr = "N"
Case eTypSim_Oth: TypSimToChr = "O"
Case eTypSim_Str: TypSimToChr = "S"
Case Else: TypSimToChr = "?"
End Select
End Function

Function VarToSimTy(pV) As eTypSim
VarToSimTy = VbTyToSim(VarType(pV))
End Function

Function VbTyToSim(pVbVarTyp As VbVarType) As eTypSim
Select Case pVbVarTyp
Case VBA.VbVarType.vbByte _
    , VBA.VbVarType.vbCurrency _
    , VBA.VbVarType.vbDecimal _
    , VBA.VbVarType.vbDouble _
    , VBA.VbVarType.vbInteger _
    , VBA.VbVarType.vbLong _
    , VBA.VbVarType.vbSingle
                                    VbTyToSim = eTypSim_Num
Case VBA.VbVarType.vbString
                                    VbTyToSim = eTypSim_Str
Case VBA.VbVarType.vbBoolean
                                    VbTyToSim = eTypSim_Bool
Case VBA.VbVarType.vbDate
                                    VbTyToSim = eTypSim_Dte
Case Else
                                    VbTyToSim = eTypSim_Oth
End Select
End Function

