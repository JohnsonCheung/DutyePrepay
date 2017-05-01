Attribute VB_Name = "nDao_Fld"
Option Compare Database
Option Explicit

Function FldNew(OFld As DAO.Field, pNmFld$, pTyp As DAO.DataTypeEnum _
    , Optional pSiz As Byte = 0 _
    , Optional pIsAuto As Boolean _
    , Optional pAlwZerLen As Boolean _
    , Optional pIsReq As Boolean _
    , Optional pDftVal$ _
    , Optional pFmtTxt$ = "" _
    , Optional pVdtTxt$ = "" _
    , Optional pVdtRul$ = "" _
    ) As Boolean
Const cSub$ = "FldNew"
Set OFld = New DAO.Field
On Error GoTo R
With OFld
    .Name = Rmv_SqBkt(pNmFld)
    If pTyp = 0 Then ss.A 1, "pTyp cannot be zero"
    .Type = pTyp
    'If pTyp = dbMemo Then Stop
    If pSiz > 0 Then .Size = pSiz
    If .AllowZeroLength <> pAlwZerLen Then .AllowZeroLength = pAlwZerLen
    If pDftVal <> "" Then
        If pTyp = dbText Then
            .DefaultValue = Q_S(pDftVal, """")
        Else
            .DefaultValue = pDftVal
        End If
    End If
    .Required = pIsReq
    If pFmtTxt <> "" Then OFld.Properties.Append OFld.CreateProperty("Format", DAO.DataTypeEnum.dbText, pFmtTxt)
    If pVdtTxt <> "" Then .ValidationText = pVdtTxt
    If pVdtRul <> "" Then .ValidationRule = pVdtRul
    If pIsAuto Then .Attributes = DAO.FieldAttributeEnum.dbAutoIncrField
End With
Exit Function
R: ss.R
E:
End Function

Function FldNew_FmRsTblF(OFld As DAO.Field, pRsTblF As DAO.Recordset) As Boolean
'     #TblF: NmFld,DaoTy,FldLen,FmtTxt,IsReq,IsAlwZerLen,DftVal,VdtTxt,VdtRul
Const cSub$ = "FldNew_FmRsTblF"
On Error GoTo R
Dim mNmFld$, mTyp As DAO.DataTypeEnum, mSiz As Byte, mIsAuto As Boolean, mIsReq As Boolean, mAlwZerLen As Boolean, mDftVal$, mFmtTxt$, mVdtTxt$, mVdtRul$
With pRsTblF
    mNmFld = !NmFld
    mTyp = !DaoTy
    mSiz = Nz(!FldLen, 0)
    mFmtTxt = Nz(!FmtTxt, "")
    mIsReq = !IsReq
    mDftVal = Nz(!DftVal.Value, "")
    Select Case mTyp
    Case DAO.DataTypeEnum.dbText, DAO.DataTypeEnum.dbMemo
        mAlwZerLen = !IsAlwZerLen
    Case Else
        mAlwZerLen = False
    End Select
    mVdtTxt = Nz(!VdtTxt, "")
    mVdtRul = Nz(!VdtRul, "")
End With
If FldNew(OFld, mNmFld, mTyp, mSiz, mIsAuto, mAlwZerLen, mIsReq, mDftVal, mFmtTxt, mVdtTxt, mVdtRul) Then ss.A 3: GoTo E
Exit Function
R: ss.R
E:
End Function

Function FldNew_FmRsTblF__Tst()
End Function

Function FldRelToStr$(A As DAO.Field)
On Error GoTo R
With A
    If .Name = .ForeignName Then
        FldRelToStr = .Name
    Else
        FldRelToStr = .Name & "=" & .ForeignName
    End If
End With
Exit Function
R: FldRelToStr = ErStr("FldRelToStr")
End Function

Function FldToDclStr$(Fld As DAO.Field)
Dim O$
Select Case Fld.Type
Case DAO.DataTypeEnum.dbChar _
   , DAO.DataTypeEnum.dbText: O = "Text(" & Fld.Size & ")"
Case Else
                              O = DaoTySql(Fld.Type)
End Select
FldToDclStr = O
End Function

Function FldToStr$(A As DAO.Field, Optional InclTy As Boolean, Optional InclVal As Boolean)
With A
    If InclTy Then
        If InclVal Then FldToStr = .Name & ":" & FldTyToStr(A) & "=" & FldToStrVal(A): Exit Function
        FldToStr = .Name & ":" & FldTyToStr(A)
        Exit Function
    End If
    If InclVal Then FldToStr = .Name & "=" & Nz(.ValidateOnSet, "Null"): Exit Function
    FldToStr = .Name
End With
End Function

Function FldToStr_Dcl$(A As DAO.Field)
FldToStr_Dcl = A.Name & " " & Cv_Fld2Dcl(A)
End Function

Function FldToStrVal$(A As DAO.Field)
On Error GoTo R
FldToStrVal = A.Value
Exit Function
R: FldToStrVal = ErStr("FldToStrVal")
End Function

Function FldTyToStr$(A As DAO.Field)
With A
    If .Type = dbText Then
        FldTyToStr = DaoTyToStr(.Type) & .Size
    Else
        FldTyToStr = DaoTyToStr(.Type)
    End If
End With
End Function

