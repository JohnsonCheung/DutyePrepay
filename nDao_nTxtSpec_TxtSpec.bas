Attribute VB_Name = "nDao_nTxtSpec_TxtSpec"
Option Compare Database
Option Explicit

Function TxtSpecCrt_Delimi(pNmSpec$, pAmFld() As tMap, Optional A As database) As Boolean
'Aim: Delete and Add one record to MSysIMEXSpecs & N records to MSysIMEXColumns to create a "text" file link spec
'     MSysIMEXSpecs  : DateDelim,DateFourDigitYear,DateLeadingZeros,DateOrder,DecimalPoint,FieldSeparator,FileType,SpecID,SpecName,SpecType,StartRow,TextDelim,TimeDelim
'     MSysIMEXColumns: Attributes,DataType,FieldName,IndexType,SkipColumn,SpecID,Start,Width
'     TxtSpec is in Am format <NmFld>=<Spec>;^^^
'     Note: <Spec>:TEXT NN,CURRENCY,LONG,INT,BYTE,DATE,SINGLE,DOUBLE,MEMO,YESNO
'           YesNo    always len=1
'           DateTime always len=8 + 1 + 6
'Hdr
'DateDelim   /
'DateFourDigitYear True
'DateLeadingZeros False
'DateOrder 5
'DecimalPoint    .
'FieldSeparator  ,
'FileType -536
'SpecID 3
'SpecName aa
'SpecType 1
'StartRow 1
'TextDelim ""
'TimeDelim:
'Det
'Attributes  0   0
'DataType    10  10
'FieldName   Obj NmObj
'IndexType   0   0
'SkipColumn  FALSE   FALSE
'SpecID  3   3
'Start   1   5
'Width   4   7

Const cSub$ = "TxtSpecCrt_Delimi"
If Dlt_TxtSpec(pNmSpec, A) Then ss.A 1: GoTo E 'Create the Spec Tables if not exist

'Create one record in MSysIMEXSpecs
Dim mSql$: mSql = Fmt_Str( _
"Insert into MSysIMEXSpecs (DateDelim,DateFourDigitYear,DateLeadingZeros,DateOrder,DecimalPoint,FieldSeparator,FileType,SpecName,SpecType,StartRow,TextDelim,TimeDelim) values " & _
                          "('/'      ,True             ,Yes             ,5        ,'.'         ,','           ,-536    ,'{0}'   ,1       ,1       ,'""'     ,':')", pNmSpec)
If Run_Sql_ByDbExec(mSql, A) Then ss.A 2: GoTo E

'Get SpecId by SpecName
Dim mSpecId&: If Fnd_ValFmSql(mSpecId, "Select SpecId from MSysIMEXSpecs where SpecName='" & pNmSpec & CtSngQ, A) Then ss.A 1: GoTo E

'Attributes
'    DataType
'        FieldName   IndexType
'                        SkipColumn
'                            SpecID
'                                Start
'                                    Width
'0   3   INT         0   0   6   1   3
'0   8   DATETIME    0   0   6   4   15
'0   5   CUR         0   0   6   19  10
'0   12  MEMO        0   0   6   29  10
'0   4   LONG        0   0   6   39  10
'0   2   BYTE        0   0   6   49  3
'0   1   YESNO       0   0   6   52  10
'0   7   DOUBLE      0   0   6   62  10
'0   10  TEXT        0   0   6   72  10
'0   6   SINGLE      0   0   6   82  10
'TEXT NN,CURRENCY,LONG,INT,BYTE,DATE,SINGLE,DOUBLE,MEMO,YESNO
'Create N records to MSysIMEXColumns
Dim J%, mStart%, mWidth%
mStart = 1: mWidth = 0
For J = 0 To Siz_Am(pAmFld) - 1
    With pAmFld(J)
        Dim mDtaTyp As Byte
        Do
            Select Case .F2
            Case "YesNo":    mDtaTyp = 1: mWidth = 1
            Case "Date": mDtaTyp = 8: mWidth = 8 + 1 + 6
            Case Else
                Dim mA$, mP%: mP = InStr(.F2, " ")
                If mP > 0 Then mA = Left(.F2, mP - 1) Else mA = .F2
                Select Case Trim(mA)
                Case "INT": mDtaTyp = 3:  mWidth = 1
                Case "CURRENCY": mDtaTyp = 5:  mWidth = 1
                Case "MEMO":     mDtaTyp = 12: mWidth = 1
                Case "LONG":     mDtaTyp = 4:  mWidth = 1
                Case "BYTE":     mDtaTyp = 2:  mWidth = 1
                Case "DOUBLE":   mDtaTyp = 7:  mWidth = 1
                Case "TEXT":     mDtaTyp = 10: mWidth = 1
                Case "SINGLE":   mDtaTyp = 6:  mWidth = 1
                Case Else
                    ss.A 4, "Invalid TypFld", eRunTimErr, "NmFld,FldSpec,Valid Spec", .F1, .F2, "TEXT NN,CURRENCY,LONG,INT,BYTE,DATE,SINGLE,DOUBLE,MEMO,YESNO": GoTo E
                End Select
            End Select
        Loop Until True

        mSql = Fmt_Str( _
        "Insert into MSysIMEXColumns (Attributes,DataType,FieldName,IndexType,SkipColumn,SpecID,Start,Width) values " & _
                                    "(0         ,{0}     ,'{1}'    ,0        ,0         ,{2}   ,{3}  ,{4})", _
                                    mDtaTyp, .F1, mSpecId, mStart, mWidth)
    End With
    mStart = mStart + mWidth
    If Run_Sql_ByDbExec(mSql, A) Then ss.A 5: GoTo E
Next
Exit Function
R: ss.R
E: TxtSpecCrt_Delimi = True: ss.B cSub, cMod, "pNmSpec,pAmFld,A", pNmSpec, ToStr_Am(pAmFld), ToStr_Db(A)
End Function

Function TxtSpecCrt_Delimi__Tst()
Const cSub$ = "TxtSpecCrt_Delimi_Tst"
Dim mWb As Workbook: If Opn_Wb_R(mWb, "P:\AppDef_Meta\MetaDb.xls") Then ss.A 1: GoTo E
Dim mAnWs$(): If Fnd_AnWs_BySetWs(mAnWs, mWb) Then ss.A 2: GoTo E
Dim J%, N%: N = Sz(mAnWs)
Dim mXls As Excel.Application: Set mXls = mWb.Application: mXls.DisplayAlerts = False
ReDim mAyFfn$(N - 1)
For J = 0 To N - 1
    Dim mWs As Worksheet: Set mWs = mWb.Sheets(mAnWs(J))
    Dim mAmFld() As tMap, mAyCno() As Byte: If Fnd_AyCnoImpFld(mAyCno, mAmFld, mWs.Range("A5")) Then ss.A 1: GoTo E

    If Clr_ImpWs(mWs.Range("A5")) Then ss.A 1: GoTo E
    'Save to Csv
    mAyFfn(J) = "c:\tmp\Exp_Ws2Tbl" & Fct.TimStmp & "_" & mWs.Name & ".csv"
    mWs.SaveAs mAyFfn(J), Excel.XlFileFormat.xlCSVWindows
    If TxtSpecCrt_Delimi(">" & mAnWs(J), mAmFld) Then ss.A 1: GoTo E
Next
Cls_Wb mWb, False, True
If Dlt_Tbl_ByPfx(">") Then ss.A 2: GoTo E
For J = 0 To N - 1
    DoCmd.TransferText acImportDelim, ">" & mAnWs(J), ">" & mAnWs(J), mAyFfn(J), True
Next
GoTo X
E:
X: mXls.DisplayAlerts = True
   Cls_Wb mWb, False, True
End Function

Function TxtSpecCrt_Fix(pNmSpec$, pLmTxtSpec$, Optional A As database) As Boolean
'Aim: Create (over-write}a Fixed len txt spec {pNmSpec} in {A} by {pLmTxtSpec}
'     Txt Spec are 2 tables definition: Delete and Add one record to MSysIMEXSpecs & N records to MSysIMEXColumns to create a "text" file link spec
'     MSysIMEXSpecs  : DateDelim,DateFourDigitYear,DateLeadingZeros,DateOrder,DecimalPoint,FieldSeparator,FileType,SpecID,SpecName,SpecType,StartRow,TextDelim,TimeDelim
'     MSysIMEXColumns: Attributes,DataType,FieldName,IndexType,SkipColumn,SpecID,Start,Width
'     TxtSpec is in Lm format <NmFld>=<Spec>;
'       <Spec>=Txt<n> Byt<n> Int<n> Sng<n> Dbl<n> Cur<n> Mem<n> YesNo DateTime
'           YesNo    always len=1
'           DateTime always len=8 + 1 + 6
Const cSub$ = "TxtSpecCrt_Fix"
If Dlt_TxtSpec(pNmSpec, A) Then ss.A 1: GoTo E

'Break pLmTxtSpec
Dim mAm() As tMap: mAm = Get_Am_ByLm(pLmTxtSpec)

'Create one record in MSysIMEXSpecs
Dim mSql$: mSql = Fmt_Str( _
"Insert into MSysIMEXSpecs (DateDelim,DateFourDigitYear,DateLeadingZeros,DateOrder,DecimalPoint,FieldSeparator,FileType,SpecName,SpecType,StartRow,TextDelim,TimeDelim) values " & _
                          "(''       ,True             ,Yes             ,5        ,'.'         ,','           ,20127   ,'{0}'   ,2       ,0       ,''       ,'')", pNmSpec)
If Run_Sql(mSql) Then ss.A 2: GoTo E

'Get SpecId by SpecName
Dim mSpecId&: If Fnd_ValFmSql(mSpecId, "Select SpecId from MSysIMEXSpecs where SpecName='" & pNmSpec & CtSngQ) Then ss.A 1: GoTo E

'Attributes
'    DataType
'        FieldName   IndexType
'                        SkipColumn
'                            SpecID
'                                Start
'                                    Width
'0   3   INT         0   0   6   1   3
'0   8   DATETIME    0   0   6   4   15
'0   5   CUR         0   0   6   19  10
'0   12  MEMO        0   0   6   29  10
'0   4   LONG        0   0   6   39  10
'0   2   BYTE        0   0   6   49  3
'0   1   YESNO       0   0   6   52  10
'0   7   DOUBLE      0   0   6   62  10
'0   10  TEXT        0   0   6   72  10
'0   6   SINGLE      0   0   6   82  10
'Txt<n> Byt<n> Int<n> Sng<n> Dbl<n> Cur<n> Mem<n> YesNo DateTime

'Create N records to MSysIMEXColumns
Dim J%, mStart%, mWidth%
mStart = 1: mWidth = 0
For J = 0 To Siz_Am(mAm) - 1
    With mAm(J)
        Dim mDtaTyp As Byte
        Do
            Select Case .F2
            Case "YesNo":    mDtaTyp = 1: mWidth = 1
            Case "DateTime": mDtaTyp = 8: mWidth = 8 + 1 + 6
            Case Else
                Dim mA$: mA = Left(.F2, 3)
                If Len(.F2) <= 3 Then ss.A 3, "Invalid data type", eRunTimErr, "NmFld,FldSpec,Valid Spec", .F1, .F2, "Txt<n> Byt<n> Int<n> Sng<n> Dbl<n> Cur<n> Mem<n> YesNo DateTime": GoTo E
                Select Case mA
                Case "INT": mDtaTyp = 3:  mWidth = Mid(.F2, 4)
                Case "CUR": mDtaTyp = 5:  mWidth = Mid(.F2, 4)
                Case "MEM": mDtaTyp = 12: mWidth = Mid(.F2, 4)
                Case "LNG": mDtaTyp = 4:  mWidth = Mid(.F2, 4)
                Case "BYT": mDtaTyp = 2:  mWidth = Mid(.F2, 4)
                Case "DBL": mDtaTyp = 7:  mWidth = Mid(.F2, 4)
                Case "TXT": mDtaTyp = 10: mWidth = Mid(.F2, 4)
                Case "SNG": mDtaTyp = 6:  mWidth = Mid(.F2, 4)
                Case Else
                    ss.A 4, "Invalid data type", eRunTimErr, "NmFld,FldSpec,Valid Spec", .F1, .F2, "Txt<n> Byt<n> Int<n> Sng<n> Dbl<n> Cur<n> Mem<n> YesNo DateTime": GoTo E
                End Select
            End Select
        Loop Until True

        mSql = Fmt_Str( _
        "Insert into MSysIMEXColumns (Attributes,DataType,FieldName,IndexType,SkipColumn,SpecID,Start,Width) values " & _
                                    "(0         ,{0}     ,'{1}'    ,0        ,0         ,{2}   ,{3}  ,{4})", _
                                    mDtaTyp, .F1, mSpecId, mStart, mWidth)
    End With
    mStart = mStart + mWidth
    If Run_Sql(mSql) Then ss.A 5: GoTo E
Next
Exit Function
R: ss.R
E: TxtSpecCrt_Fix = True: ss.B cSub, cMod, "pNmSpec,pLmTxtSpec,A", pNmSpec, pLmTxtSpec, ToStr_Db(A)
End Function

Function TxtSpecCrt_Fix__Tst()
If TxtSpecCrt_Fix("A2Test", "I=Int3, A=Txt1, B=Txt2, C=Txt3") Then Stop: GoTo E
Stop
Dim mF As Byte: If Opn_Fil_ForOutput(mF, "c:\tmp\aa.txt", True) Then Stop: GoTo E
Print #mF, "123XAA 22"
Print #mF, "12 YAB  2"
Print #mF, "1  ZAB   "
Print #mF, "123 AB222"
Close #mF
DoCmd.TransferText acImportFixed, "A2Test", "#Tmp", "c:\tmp\aa.txt", False
DoCmd.OpenTable "#Tmp"
Stop
GoTo X
E:
X: Close mF
End Function

